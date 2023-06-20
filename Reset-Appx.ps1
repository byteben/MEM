<#
.SYNOPSIS
    Remove a built-in modern app from Windows, for All Users, and reinstall using WinGet
    
.DESCRIPTION
    This script will remove a specific built-in AppXPackage, for All Users, and also the AppXProvisionedPackage if it exists
    When deploying apps from the new store, via Intune, in the SYSTEM context, an error appears for the deployment if the same app was previous deployed in the USER context
    "The application was not detected after installation completed successfully (0x87D1041C)"
    This script will remove all existing instances of the app so the app from Intune can be installed sucessfully
    AppxPackage removal can fail if the app was installed from the Microsoft Store. This script will re-register the app for All Users in that instance to allow for removal
    WinGet will then install the app in the scope of the machine

    .NOTES
    FileName:       Reset-Appx.ps1
    Created:        12th June 2023
    Updated:        18th June 2023
    Author:         Ben Whitmore @ PatchMyPC (Thanks to Bryan Dam @bdam555 for assisted research and blog at https://patchtuesday.com/blog/intune-microsoft-store-integration-app-migration-failure/)
    Contact:        @byteben
    Manifest:       Company Portal manifest: https://storeedgefd.dsx.mp.microsoft.com/v9.0/packageManifests/9WZDNCRFJ3PZ
    
    Version History:

    1.06.18.0 - Bug Fixes and New Function

    -   New function "Test-WinGet" added to test if WinGet is installed and working. AppXPackages will not be removed if there is an issue with the WinGet command line
        -   Tests WinGet package is installed
        -   Tests WinGet.exe is working
        -   Tests if WinGet command line failure occurs because Visual C++ 14.x Redistributable is not installed
    -   Fixed evaluation AppXProvisionedPackage results
    -   Minor log and output bugs fixed

    1.06.15.0 - Bug Fixes

        -   Fixed an issue where AppxProvisionedPackage removal was not working as expected
        -   Fixed an issue where winGetPath would return more than one package. We now sort results in descending order and select the first 1
        -   Fixed an issue where --accept-source-agreements was omitted from "winget list" to search for an existing app

    1.06.14.0 - Minor Fixes

        -   Fixed an issue with the WinGet output array object and now we loop through each array item looking for output strings
        -   Fixed an issue where AppxProvisionedPackage would indicate succesful removal even if it wasn't installed
        -   Fixed an issue where WinGet install and list commands returned unexecpeted result. Now using Package ID instead of Package Name
        -   Fixed an issue where "WingGet list" would not detect the app after it was installed. Using Get-AppxPackage instead to improve reliability of detection
        -   Changed path tests for WinGet binary
        -   Increased logging detail for each test

    1.06.12.0 - Release
    
.PARAMETER removeApp
    Specify the AppxPackage and AppxProivisionedPackage to remove
    The parameter is defined at the top of the script so it can be used as an Intune Script (which does not accept params)

.PARAMETER winGetAppInstall
    Boolean True or False. Should an attempt be made to reinstall the app, with WinGet, after it has been removed

.PARAMETER winGetApp
    Specify the app id to reinstall using WinGet. Use Winget Search "*appname*" to understand which id you should use

.PARAMETER winGetAppName
    Specify the app to reinstall using WinGet. Use Winget Search "*appname*" to understand which name you should use

.PARAMETER winGetAppSource
    Specify WinGet source to use. Typically this will be msstore for apps with the issue outlined in the description of this script

.EXAMPLE
    .\Reset-Appx.ps1

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$removeApp = 'Microsoft.CompanyPortal',
    [bool]$winGetAppInstall = $true,
    [string]$winGetApp = '9WZDNCRFJ3PZ',
    [string]$winGetAppName = 'Company Portal',
    [string]$winGetAppSource = 'msstore',
    [string]$logID = 'Main'
)

Begin {

    If (([Security.Principal.WindowsIdentity]::GetCurrent()).IsSystem -eq $false) {
        Write-Error 'This script needs to run as SYSTEM'
        break
    }

    # Create variables
    $removeAppxPackage = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } -ErrorAction SilentlyContinue
    $removeAppxProvisionedPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $removeApp } | Select-Object -ExpandProperty PackageName -ErrorAction SilentlyContinue
    $global:winGetPath = $null
}

Process {
    
    # Functions
    function Write-LogEntry {
        param(
            [parameter(Mandatory = $true)]
            [string[]]$logEntry,
            [string]$logID,
            [parameter(Mandatory = $false)]
            [string]$logFile = "$($env:ProgramData)\Microsoft\IntuneManagementExtension\Logs\Reset-Appx.log",
            [ValidateSet(1, 2, 3, 4)]
            [string]$severity = 4,
            [string]$component = [string]::Format('{0}:{1}', $logID, $($MyInvocation.ScriptLineNumber))
        )

        Begin {
            $dateTime = Get-Date
            $date = $dateTime.ToString("MM-dd-yyyy", [Globalization.CultureInfo]::InvariantCulture)
            $time = $dateTime.ToString("HH:mm:ss.ffffff", [Globalization.CultureInfo]::InvariantCulture)
        }

        Process {
            
            try {
                # Extract log object and construct format for log line entry
                foreach ($log in $logEntry) {
                    $logDetail = [string]::Format('<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">', $log, $time, $date, $component, $context, $severity, $PID)

                    # Attempt log write
                    try {
                        $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $logFile, 'Append'
                        $streamWriter.WriteLine($logDetail)
                        $streamWriter.Close()
                    }
                    catch {
                        Write-Error -Message "Unable to append log entry to $logFile file. Error message: $($_.Exception.Message)"
                    }
                }
            }
            catch [System.Exception] {
                Write-Warning -Message "Unable to append log entry to $($fileName) file"
            }
        }
    }

    function Remove-AppxPkg {
        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name
        )

        # Attempt to remove AppxPackage
        Write-Host "Processing AppxPackage: $($removeApp)"
        Write-LogEntry -logEntry "Processing appx package: $($removeApp)" -logID $logID 

        if (-not[string]::IsNullOrEmpty($removeAppxPackage)) {
           
            try {
                
                # List users with the AppxPackage installed
                Write-Host "Found the following users with AppxPackage '$($removeApp)':"
                Write-LogEntry -logEntry "Found the following users with AppxPackage '$($removeApp)':" -logID $logID 
                
                foreach ($removeAppxPackageUserInfo in $removeAppxPackage.PackageUserInformation) {
                    $removeAppxPackageUser = ( $removeAppxPackageUserInfo | Select-Object -ExpandProperty UserSecurityId).UserName
                    Write-Host "User: '$($removeAppxPackageUser)', InstallState: '$($removeAppxPackageUserInfo.InstallState)'"
                    Write-LogEntry -logEntry "User: '$($removeAppxPackageUser)', InstallState: '$($removeAppxPackageUserInfo.InstallState)'" -logID $logID 
                }

                Write-Host "Removing AppxPackage: $($removeApp)"
                Write-LogEntry -logEntry "Removing AppxPackage: $($removeApp)" -logID $logID  
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            catch [System.Exception] {
                
                if ( $_.Exception.Message -like "*HRESULT: 0x80073CF1*") {
                    Write-Warning "AppxPackage removal failed. Error: 0x80073CF1. The manifest for the '$($removeApp)' needs to be re-registered before it can be removed."
                    Write-LogEntry -logEntry "AppxPackage removal failed. Error: 0x80073CF1. The manifest for the '$($removeApp)' needs to be re-registered before it can be removed." -logID $logID 
                    $removeAppxPackageError0x80073CF1 = $true
                }
                elseif ($_.Exception.Message -like "*failed with error 0x80070002*") {
                    Write-Warning "AppxPackage removal failed. Error 0x80070002"
                    Write-LogEntry -logEntry "AppxPackage removal failed. Error 0x80070002" -logID $logID 
                    $removeAppxPackageError0x80070002 = $true
                }
                else {
                    Write-Warning -Message "Removing AppxPackage '$($removeApp)' failed"
                    Write-Warning -Message $_.Exception.Message
                    Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)' failed" -logID $logID -severity 3
                    Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3
                }
            }
            
            # Test removal was successful
            $testAppx = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp }

            if ([string]::IsNullOrEmpty($testAppx)) {
                Write-Host "All instances of AppxPackage: $($removeApp) were removed succesfully"
                Write-LogEntry -logEntry "All instances of AppxPackage: $($removeApp) were removed succesfully" -logID $logID  
            }
            else {
                Write-Warning -Message "Removing AppxPackage '$($removeApp)' for all users was not succesful"
                Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)' for all users was not succesful" -logID $logID -severity 3
            }
        }
        else {
            Write-Output "Did not attempt removal of the AppxPackage '$($removeApp)' because it was not found"
            Write-LogEntry -logEntry "Did not attempt removal of the AppxPackage '$($removeApp)' because it was not found" -logID $logID -severity 2
        }

        # Re-register AppxPackage for all users and attempt removal again
        if ( $removeAppxPackageError0x80073CF1 ) {
            Write-Host "Attempting to re-register AppxPackage '$($removeApp)'..."
            Write-LogEntry -logEntry "Attempting to re-register AppxPackage '$($removeApp)'..." -logID $logID 

            try {
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }
                
                If (-not[string]::IsNullOrEmpty( { Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } } )) {
                    Write-Host "AppxPackage '$($removeApp)' registered succesfully"
                    Write-LogEntry -logEntry "AppxPackage '$($removeApp)' registered succesfully" -logID $logID 
                }
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            catch {
                Write-Warning -Message "Re-registering AppxPackage '$($removeApp)' failed: $($_.Exception.Message)"
                Write-LogEntry -logEntry "Re-registering AppxPackage '$($removeApp)' failed: $($_.Exception.Message)" -logID $logID -severity 3
            }
        }
    }

    function Remove-AppxProvPkg {
        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name
        )
    
        # Attempt to remove AppxProvisionedPackage
        If (-not[string]::IsNullOrEmpty($removeAppxProvisionedPackageName)) {

            try {
                Write-Host "Removing AppxProvisioningPackage: $($removeAppxProvisionedPackageName)"
                Write-LogEntry -logEntry "Removing AppxProvisioningPackage: $($removeAppxProvisionedPackageName)" -logID $logID 
                Write-Host "Get-AppxProvisionedPackage -Online | Where-Object { `$_.PackageName -eq $($removeAppxProvisionedPackageName) } | Remove-AppxProvisionedPackage -AllUsers -ErrorAction Stop"
                Write-LogEntry -logEntry "Get-AppxProvisionedPackage -Online | Where-Object { `$_.PackageName -eq $($removeAppxProvisionedPackageName) } | Remove-AppxProvisionedPackage -AllUsers -ErrorAction Stop" -logID $logID 

                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $removeAppxProvisionedPackageName } | Remove-AppxProvisionedPackage -AllUsers -ErrorAction Stop
                $removeAppxProvisionedPackageRemovalAttempt = $true
            }
            catch [System.Exception] {
                Write-Warning -message "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed"
                Write-Warning -message $_.Exception.Message
                Write-LogEntry -logEntry "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed" -logID $logID -severity 3
                Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3
            }
        }
        else {
            Write-Output "Did not attempt removal of the AppxProvisionedPackage '$($removeApp)' because it was not found"
            Write-LogEntry -logEntry "Did not attempt removal of the AppxProvisionedPackage '$($removeApp)' because it was not found" -logID $logID -severity 2
        }

        # Test removal was successful
        if ($removeAppxProvisionedPackageRemovalAttempt -eq $true) {

            $testAppxProv = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $removeAppxProvisionedPackageName }

            if ([string]::IsNullOrEmpty($testAppxProv)) {
                Write-Host "AppxProvisionedPackage: $($removeApp) was removed succesfully"
                Write-LogEntry -logEntry "AppxProvisionedPackage: $($removeApp) was removed succesfully" -logID $logID  
            }
            else {
                Write-Warning -Message "AppxProvisionedPackage: $($removeApp) removal was unsuccessful"
                Write-LogEntry -logEntry "AppxProvisionedPackage: $($removeApp) removal was unsuccessful" -logID $logID -severity 3
            }
        }
    }

    function Test-WinGet {

        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [string]$winGetPackageName = 'Microsoft.DesktopAppInstaller',
            [string]$winGetBinary = 'winget.exe'
        )

        # Test the WinGet package and other dependancies are installed
        Write-Host "Testing the WinGet package and other dependancies are installed"
        Write-LogEntry -logEntry "Testing the WinGet package and other dependancies are installed" -logID $logID

        try {
            $global:winGetPath = (Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq $winGetPackageName }).InstallLocation | Sort-Object -Descending | Select-Object -First 1
        }
        catch {
            $testWinGetFail = $true
            Write-Warning -Message "There was a problem getting details of the '$($winGetPackageName)' package"
            Write-LogEntry -logEntry "There was a problem getting details of the '$($winGetPackageName)' package" -logID $logID -severity 3
        }

        if ([string]::IsNullOrEmpty($global:winGetPath)) {
            $testWinGetFail = $true
            Write-Warning "The '$($winGetPackageName)' package was not found'"
            Write-LogEntry -logEntry "The '$($winGetPackageName)' package was not found" -logID $logID  -severity 3
        }
        else {
            $winGetBinaryPath = Join-Path -Path $global:winGetPath -ChildPath 'WinGet.exe'

            try {
                
                if (Test-Path -Path $winGetBinaryPath ) {
                    Write-Host "The '$($winGetBinary)' binary was found at '$($winGetBinaryPath)'"
                    Write-LogEntry -logEntry "The '$($winGetBinary)' package was found at '$($winGetBinaryPath)'" -logID $logID
                }
                else {
                    $testWinGetFail = $true
                    Write-Warning "The '$($winGetPackageName)' package was found at '$($global:winGetPath)' but the WinGet binary was not found at '$($winGetBinaryPath)'"
                    Write-LogEntry -logEntry "The '$($winGetPackageName)' package was found at '$($global:winGetPath)' but the WinGet binary was not found at '$($winGetBinaryPath)'" -logID $logID -severity 3
                }
            }
            catch {
                $testWinGetFail = $true
                Write-Warning "An error was encounted trying to validate the path to WinGet.exe at '$($winGetBinaryPath)'"
                Write-LogEntry -logEntry "An error was encounted trying to validate the path to WinGet.exe at '$($winGetBinaryPath)'" -logID $logID -severity 3
            }
        }
        
        if ($testWinGetFail) {
            Write-Warning "The '$($winGetPackageName)' package was not found or the WinGet binary was not found at '$($winGetBinaryPath)'. Cannot continue"
            Write-LogEntry -logEntry "The '$($winGetPackageName)' package was not found or the WinGet binary was not found at '$($winGetBinaryPath)'. Cannot continue" -logID $logID -severity 3
        }
        else {
            
            # Test WinGet running as SYSTEM
            try {
                Set-Location $global:winGetPath
                $winGetTest = .\winget.exe --version

                if (-not[string]::IsNullOrEmpty($winGetTest)) {
                    Write-Host "The WinGet binary was validated"
                    Write-Host "WinGet version is '$($winGetTest)'"
                    Write-LogEntry -logEntry "The WinGet binary was validated" -logID $logID
                    Write-LogEntry -logEntry "WinGet version is '$($winGetTest)'" -logID $logID
                }
                else {
                    $testWinGetFail = $true
                    Write-Warning "No output was detected while testing the WinGet version"
                    Write-Warning "WinGet has a dependency on this VC++ redistributable 14.x when running in the SYSTEM context. Ensure VC++ redistributable 14.x or higher is installed"
                    Write-LogEntry -logEntry "No output was detected while testing the WinGet version"-logID $logID -severity 3
                    Write-LogEntry -logEntry "WinGet has a dependency on this VC++ redistributable 14.x when running in the SYSTEM context. Ensure VC++ redistributable 14.x or higher is installed" -logID $logID -severity 3
                }
            }
            catch {
                $testWinGetFail = $true
                Write-Warning "An error was encountered trying to run the WinGet binary using the command '$($winGetTestExpression)'"
                Write-LogEntry -logEntry "An error was encountered trying to run the WinGet binary using the command '$($winGetTestExpression)'" -logID $logID -severity 3
            }
        }

        if ($testWinGetFail) {
            return 'Failed'
        }
        else {
            return 'Passed'
        }
    }

    function Install-WinGetApp {
        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name
        )
    
        # Attempt to install app using WinGet
        try {

            Write-Host "Checking if '$($winGetAppName)' is installed using Id '$($winGetApp)'..."
            Write-Host "winget.exe list --id $($winGetApp) --source $($winGetAppSource) --accept-source-agreements"
            Write-LogEntry -logEntry "Checking if '$($winGetAppName)' is installed using Id '$($winGetApp)'..." -logID $logID
            Write-LogEntry -logEntry "winget.exe list --id '$($winGetApp)' --source $($winGetAppSource) --accept-source-agreements" -logID $logID 

            Set-Location $global:winGetPath
            $winGetTest = .\winget.exe list --id $winGetApp --source $winGetAppSource --accept-source-agreements
                
            foreach ($line in $winGetTest) {
                
                if ($line -like "*No installed package found*") {
                    $winGetAppMissing = $true
                }

                if ($line -like $winGetApp) {
                    $winGetAppAlreadyInstalled = $true
                }
            }

            if ($winGetAppMissing -eq $true) {

                Write-Host "The 'Winget list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetApp)', was not installed. Installing '$($winGetAppName)' using WinGet command line..."
                Write-LogEntry -logEntry "The 'Winget list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetApp)', was not installed. Installing '$($winGetAppName)' using WinGet command line..." -logID $logID 

                try {
                    Write-Host "Installing '$($winGetAppName)', with Id '$($winGetApp)', using the WinGet command line"
                    Write-LogEntry -logEntry "Installing '$($winGetAppName)', with Id '$($winGetApp)', using the WinGet command line" -logID $logID
                    Write-Host ".\winget.exe install --id '$winGetApp' --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine"
                    Write-LogEntry -logEntry ".\winget.exe install --Id '$winGetApp' --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine" -logID $logID

                    .\winget.exe install --id $winGetApp --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine
                    $winGetAppInstallAttempted = $true
                }
                catch {
                    Write-Warning -Message "There was an error installing '$($winGetAppName)', with Id '$($winGetApp)', using the WinGet command line"
                    Write-Warning -Message "$($_.Exception.Message)"
                    Write-LogEntry -logEntry "There was an error installing '$($winGetAppName)', with Id '$($winGetApp)', using the WinGet command line" -logID $logID -severity 3
                    Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3
                }
            }
                
            if ($winGetAppAlreadyInstalled) {
                Write-Host "The 'Winget list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetApp)', is already installed"
                Write-LogEntry -logEntry "The 'Winget list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetApp)', is already installed" -logID $logID -severity 2
            }
        }
        catch {
            Write-Warning -Message "Error while running the WinGet command line to check if '$($winGetAppName)' is already installed"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "Error while running the WinGet command line to check if $($winGetAppName) is already installed" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3
        }

        # Test package was succesfully installed

        Write-Host "Checking if '$($winGetAppName)' is installed using Get-AppXPackage..."
        Write-Host "Get-AppXPackage -AllUsers | Where-Object { `$_.Name -like `"*$($removeApp)*`" } -ErrorAction Stop"
        Write-LogEntry -logEntry "Checking if '$($winGetAppName)' is installed using Get-AppXPackage..." -logID $logID
        Write-LogEntry -logEntry "Get-AppXPackage -AllUsers | Where-Object { `$_.Name -like `"*$($removeApp)*`" } -ErrorAction Stop" -logID $logID 

        if ($winGetAppInstallAttempted) {
            $testWinGetInstall = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } -ErrorAction Stop

            if ($testWinGetInstall.Name -eq $removeApp) {

                Write-Host "Success: The '$($winGetAppName)' app, with Id $($winGetApp), installed succesfully. Check the Winget logs at 'C:\Windows\Temp\WinGet\defaultState' for more information"
                Write-LogEntry -logEntry "Success: The '$($winGetAppName)' app, with Id '$($winGetApp)', installed succesfully. Check the Winget logs at 'C:\Windows\Temp\WinGet\defaultState' for more information" -logID $logID
            }
            else {
                Write-Warning -Message "Error: The '$($winGetAppName)' app, with Id '$($winGetApp)', did not install succesfully. Check the Winget logs at 'C:\Windows\Temp\WinGet\defaultState' for more information"
                Write-LogEntry -logEntry "Error: The '$($winGetAppName)' app, with Id '$($winGetApp)', did not install succesfully. Check the Winget logs at 'C:\Windows\Temp\WinGet\defaultState' for more information" -logID $logID -severity 3
            }
        }
    }

    # Initial logging
    Write-Host '** Starting processing the script' 
    Write-LogEntry -logEntry '** Starting processing the script' -logID $logID 

    # Call Functions
    if ($winGetAppInstall -eq $true) {
        $winGetTestResult = Test-Winget

        if ($winGetTestResult -eq 'Passed') {
            Remove-AppxPkg
            Remove-AppxProvPkg
            Install-WinGetApp
        }
        else {
            Write-Warning "The `$winGetAppInstall paramter was set to true but WinGet tests failed. Aborting other functions"
            Write-LogEntry -logEntry "The `$winGetAppInstall paramter was set to true but WinGet tests failed. Aborting other functions" -logID $logID -severity 3
        }
    }
    else {
        Write-Host "The `$winGetAppInstall paramter was set to false. Will not attempt to re-install the package using WinGet after '$($removeApp)' is removed. Continuing with the script"
        Write-LogEntry -logEntry Write-Host "The `$winGetAppInstall paramter was set to false. Will not attempt to re-install the package using WinGet after '$($removeApp)' is removed. Continuing with the script" -logID $logID

        Remove-AppxPkg
        Remove-AppxProvPkg
    }

    # Complete
    Write-Output "Finished processing the script"
    Write-LogEntry -logEntry "Finished processing the script" -logID $logID
    Set-Location $PSScriptRoot
}
