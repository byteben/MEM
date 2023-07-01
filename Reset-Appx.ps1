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
    Updated:        1st Julye 2023
    Author:         Ben Whitmore @byteben
    Contributors:   Bryan Dam @bdam555 for assisted research and blog at https://patchtuesday.com/blog/intune-microsoft-store-integration-app-migration-failure/)
                    Adam Cook @codaamok for refactor advice and down the rabbit hole testing with WinGet
    Contact:        @byteben
    Manifest:       Company Portal manifest: https://storeedgefd.dsx.mp.microsoft.com/v9.0/packageManifests/9WZDNCRFJ3PZ
    
    Version History:

    1.07.01.0 - Bug Fixes, Enhancements and Refactoring

    -   Test-WinGetBinary returns the path to the WinGet binary if it was found with a test result passed or failed
    -   Renamed variable $winGetApp to $winGetAppId to avoid confusion with the WinGet app name
    -   Refactored functions to accept value from pipeline instaead of declaring global variables
    -   Fixed an issue where testing the if the WinGet app was installed would fail because the -like operater did not evaluate the WinGet app Id correctly
    -   New functions:- 
        -   Test-AppxProvisionedPackage
        -   Test-AppxPackage
        -   Test-WinGetBinary
        -   Install-WinGetApp
        -   Test-WinGetPath
        -   Test-AppxPackageUserInformation
        -   Remove-AppxProvPackage
        -   Register-AppxPackage

    1.06.27.0 - Bug Fixes and Enhancements

    -   Now installing app with winget using --name instead of --id to avoid an issue where winget log indicated "No app found matching input criteria"
    -   Fixed an issue where winget would provision the appxProvisionedPackage but the package manifest was not registered resulting in the app not being installed for the logged on user
    -   Replaced like with eq for consistency when searching for appxProvisionedPackages and appxPackages

    1.06.20.0 - Bug Fixes

    -   Fixed an issue where winGetPath was not declared globally. Thanks https://github.com/Sn00zEZA for reporting

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

.PARAMETER winGetAppId
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
    [string]$winGetAppId = '9WZDNCRFJ3PZ',
    [string]$winGetAppName = 'Company Portal',
    [string]$winGetAppSource = 'msstore',
    [string]$winGetPackageName = 'Microsoft.DesktopAppInstaller',
    [string]$winGetBinary = 'winget.exe',
    [int]$winGetRetries = 10,
    [int]$appxWaitTimerSeconds = 30,
    [string]$logID = 'Main'
)

Begin {

    If (([Security.Principal.WindowsIdentity]::GetCurrent()).IsSystem -eq $false) {
        Write-Error 'This script needs to run as SYSTEM'
        break
    }
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
    Function Test-AppxPackage {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeApp,
            [int]$appxWaitTimerSeconds
        )

        # Test the if the AppxPackage is installed
        Write-Host "Checking if the AppxPackage '$($removeApp)' is installed..."
        Write-LogEntry -logEntry "Checking if the AppxPackage '$($removeApp)' is installed..." -logID $logID
        Write-Host "Get-AppxPackage -AllUsers | Where-Object { `$_.Name -eq '$($removeApp)' } -ErrorAction Stop"
        Write-LogEntry -logEntry "Get-AppxPackage -AllUsers | Where-Object { `$_.Name -eq '$($removeApp)' } -ErrorAction Stop" -logID $logID

        try {

            $testAppxPackage = Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq $removeApp } -ErrorAction Stop

            if ($testAppxPackage.Name -eq $removeapp) {
                Write-Host "The '$($removeApp)' AppxPackage was found"
                Write-LogEntry -logEntry "The '$($removeApp)' AppxPackage was found" -logID $logID

                # Check if the AppxPackage is installed for the SYSTEM account and staged
                Test-appxPackageUserInformation -PackageUserInformation $testAppxPackage.PackageUserInformation -removeApp $removeApp

                try {

                    # Wait and check again to ensure the appxPackage is not cleaning up after a failed staging
                    Write-Host "Waiting '$($appxWaitTimerSeconds)' seconds to repeat the test '$($removeApp)' to ensure the appxPackage is not cleaning up after a failed staging"
                    Write-LogEntry -logEntry "Waiting '$($appxWaitTimerSeconds)' seconds to repeat the test '$($removeApp)' to ensure the appxPackage is not cleaning up after a failed staging" -logID $logID
                
                    # Timer to wait for appx staging cleanup
                    for ($t = $appxWaitTimerSeconds; $t -ge 0; $t--) {
                        Write-Host -NoNewLine "$t.."
                        Start-Sleep -Seconds 1
                    }

                    if ($t -eq 0) {
                        Write-Host "Waited '$($appxWaitTimerSeconds)' seconds for appx staging cleanup"
                        Write-LogEntry -logEntry "Waited '$($appxWaitTimerSeconds)' seconds for appx staging cleanup" -logID $logID
                    }
                    
                    $testAppxPackage = Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq $removeApp } -ErrorAction Stop

                    if ($testAppxPackage.Name -eq $removeapp) {
                        Write-Host "The '$($removeApp)' AppxPackage was found after waiting '$($appxWaitTimerSeconds)' seconds for appx staging cleanup"
                        Write-LogEntry -logEntry "The '$($removeApp)' AppxPackage was found after waiting '$($appxWaitTimerSeconds)' seconds for appx staging cleanup" -logID $logID

                        # Check if the AppxPackage is installed for the SYSTEM account and staged
                        Test-appxPackageUserInformation -PackageUserInformation $testAppxPackage.PackageUserInformation -removeApp $removeApp   
                    }
                    else {
                        Write-Warning -Message "The '$($removeApp)' AppxPackage was not found after waiting '$($appxWaitTimerSeconds)' seconds for appx staging cleanup"
                        Write-LogEntry -logEntry "The '$($removeApp)' AppxPackage was not found after waiting '$($appxWaitTimerSeconds)' seconds for appx staging cleanup" -logID $logID -severity 3

                        return @{Result = 'Not Installed'; Users = $null }
                    }
                }
                catch {
                    Write-Warning -Message "Error while running the Get-AppxPackage command line to check if '$($removeApp)' is installed"
                    Write-Warning -Message "$($_.Exception.Message)"
                    Write-LogEntry -logEntry "Error while running the Get-AppxPackage command line to check if '$($removeApp)' is installed" -logID $logID -severity 3
                    Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3

                    return @{Result = 'Fatal Error'; Users = $null }
                    $LASTEXITCODE = 1
                }
            }
            else {
                Write-Host "The '$($removeApp)' AppxPackage is not installed"
                Write-LogEntry -logEntry "The '$($removeApp)' AppxPackage is not installed" -logID $logID

                return @{Result = 'Not Installed'; Users = $null }
            }
        }
        catch {
            Write-Warning -Message "Error while running the Get-AppxPackage command line to check if '$($removeApp)' is installed"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "Error while running the Get-AppxPackage command line to check if '$($removeApp)' is installed" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3

            return @{Result = 'Fatal Error'; Users = $null }
            $LASTEXITCODE = 1
        }
    }

    Function Test-AppxPackageUserInformation {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeapp,
            [array]$packageUserInformation
        )

        # Check if the appxPackage is staged for the SYSTEM account which would indicate a failed appx install
        # Create an array to store any users found with the appxPackage installed
        $userList = @()
                        
        try {

            foreach ($user in $packageUserInformation) {

                if ($user.UserSecurityId.sid -eq 'S-1-5-18' -or $user -like "*S-1-5-18*") {
                    $sysemStagedFound = $true

                    Write-Host "The '$($removeApp)' AppxPackage is staged for the SYSTEM account which could indicate a failed appxPackage install by WinGet"
                    Write-LogEntry -logEntry "The '$($removeApp)' AppxPackage is staged for the SYSTEM account which could indicate a failed appxPackage install by WinGet" -logID $logID
                }
                
                # If an array is returned without columns regex the string and return only the username
                if ($user -match "^S-\d-(?:\d+-){1,14}\d+") {
                    $username = ($user -replace '^.*\[(.*?)\].*$', '$1')
                }
                else {
                    $username = $user.UserSecurityId.UserName
                }
                $userList += $username
            }

            if ($sysemStagedFound) {
                return @{Result = 'SYSTEM Staged'; Users = $userList }
            }
            else {
                return @{Result = 'Installed'; Users = $userList }
            } 
        }
        catch {
            Write-Warning -Message "Error while running the Test-AppxPackageUserInformation command line to check if '$($removeApp)' is installed"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "Error while running the Test-AppxPackageUserInformation command line to check if '$($removeApp)' is installed" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3

            return @{Result = 'Fatal Error'; Users = $null }
            $LASTEXITCODE = 1
        }
    }

    function Remove-AppxPkg {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeApp
        )

        # Note: Function name shortened because of clash with cmdlet Remove-AppxPackage
        # Attempt to remove AppxPackage
        try {
            Write-Host "Removing AppxPackage '$($removeApp)'..."
            Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)'..." -logID $logID 
            Write-Host "Remove-AppxPackage -AllUsers -Package '$($removeApp)' -ErrorAction Stop"
            Write-LogEntry -logEntry "Remove-AppxPackage -AllUsers -Package '$($removeApp)' -ErrorAction Stop" -logID $logID 

            Remove-AppxPackage -AllUsers -Package $removeApp -ErrorAction Stop
        }
        catch [System.Exception] {
            if ( $_.Exception.Message -like "*HRESULT: 0x80073CF1*") {
                Write-Warning "AppxPackage removal failed. Error: 0x80073CF1. The manifest for the '$($removeApp)' needs to be re-registered before it can be removed."
                Write-LogEntry -logEntry "AppxPackage removal failed. Error: 0x80073CF1. The manifest for the '$($removeApp)' needs to be re-registered before it can be removed." -logID $logID 
                    
                return @{Result = 'Failed'; Reason = '0x80073CF1' }
            }
            elseif ($_.Exception.Message -like "*failed with error 0x80070002*") {
                Write-Warning "AppxPackage removal failed. Error 0x80070002"
                Write-LogEntry -logEntry "AppxPackage removal failed. Error 0x80070002" -logID $logID 

                return @{Result = 'Failed'; Reason = '0x80070002' }
            }
            else {
                Write-Warning -Message "Removing AppxPackage '$($removeApp)' failed"
                Write-Warning -Message $_.Exception.Message
                Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)' failed" -logID $logID -severity 3
                Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3

                return @{Result = 'Failed'; Reason = 'Other' }
            }
        }
            
        # Test removal was successful
        $testAppxPackageResult = Test-AppxPackage -removeApp $removeApp -appxWaitTimerSeconds $appxWaitTimerSeconds

        if ($testAppxPackageResult.Result -eq 'Not Installed') {
            Write-Host "All instances of AppxPackage: $($removeApp) were removed succesfully"
            Write-LogEntry -logEntry "All instances of AppxPackage: $($removeApp) were removed succesfully" -logID $logID  

            return @{Result = 'Not Installed'; Reason = $null }
        }
        else {
            Write-Warning -Message "Removing AppxPackage '$($removeApp)' for all users was not succesful"
            Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)' for all users was not succesful" -logID $logID -severity 3

            Return @{Result = 'Failed'; Reason = 'Other' }
        }
    }

    Function Register-AppxPackage {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeApp,
            [string]$appxWaitTimerSeconds
        )

        # Re-register AppxPackage for all users and attempt removal again
        try {
            Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq $removeApp } | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" } -ErrorAction Stop
                
            If ((Test-AppxPackage -removeApp $removeApp -appxWaitTimerSeconds $appxWaitTimerSeconds).Result -eq 'Installed') {
                Write-Host "AppxPackage '$($removeApp)' was registered succesfully"
                Write-LogEntry -logEntry "AppxPackage '$($removeApp)' was registered succesfully" -logID $logID 

                return @{Result = 'Success' }
            }
            else {
                Write-Warning -Message "Re-registering AppxPackage '$($removeApp)' failed"
                Write-LogEntry -logEntry "Re-registering AppxPackage '$($removeApp)' failed" -logID $logID -severity 3

                return @{Result = 'Failure' }
            }
        }
        catch {
            Write-Warning -Message "Re-registering AppxPackage '$($removeApp)' failed"
            Write-Warning -Message $_.Exception.Message
            Write-LogEntry -logEntry "Re-registering AppxPackage '$($removeApp)' failed" -logID $logID -severity 3
            Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3

            return @{Result = 'Fatal Error' }
            $LASTEXITCODE = 1
        }
    }

    Function Test-AppxProvisionedPackage {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeApp
        )

        # Test the if the AppxProvisionedPackage is installed
        Write-Host "Checking if the AppxProvisionedPackage '$($removeApp)' is installed..."
        Write-Host "Get-AppxProvisionedPackage -Online | Where-Object { `$_.DisplayName -eq '$($removeApp)' } | Select-Object DisplayName, PackageName -ErrorAction Stop"
        Write-LogEntry -logEntry "Checking if the AppxProvisionedPackage '$($removeApp)' is installed..." -logID $logID
        Write-LogEntry -logEntry "Get-AppxProvisionedPackage -Online | Where-Object { `$_.DisplayName -eq '$($removeApp)' } | Select-Object DisplayName, PackageName -ErrorAction Stop" -logID $logID

        try {
            $testAppxProvisionedPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $removeApp } | Select-Object DisplayName, PackageName -ErrorAction Stop

            if ($testAppxProvisionedPackage.DisplayName -eq $removeapp) {
                Write-Host "The '$($removeApp)' AppxProvisionedPackage was found"
                Write-LogEntry -logEntry "The '$($removeApp)' AppxProvisionedPackage was found" -logID $logID
                
                return @{Result = 'Installed'; PackageName = $testAppxProvisionedPackage.PackageName }
            }
            else {
                Write-Host "The '$($removeApp)' AppxProvisionedPackage is not installed"
                Write-LogEntry -logEntry "The '$($removeApp)' AppxProvisionedPackage is not installed" -logID $logID

                return @{Result = 'Not Installed'; PackageName = $null }
            }
        }
        catch {
            Write-Warning -Message "Error while running the Get-AppxProvisionedPackage command line to check if '$($removeApp)' is installed"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "Error while running the Get-AppxProvisionedPackage command line to check if '$($removeApp)' is installed" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3

            return @{Result = 'Fatal Error'; PackageName = $null }
            $LASTEXITCODE = 1
        }
    }

    function Remove-AppxProvPackage {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$removeAppxProvisionedPackageName,
            [string]$removeApp
        )

        # Note: Function name shortened because of clash with cmdlet Remove-AppxProvisionedPackage
        # Attempt to remove AppxProvisionedPackage
        try {
            Write-Host "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)'..."
            Write-LogEntry -logEntry "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)'..." -logID $logID 
            Write-Host "Remove-AppxProvisionedPackage -Online -PackageName '$($removeAppxProvisionedPackageName)' -AllUsers -ErrorAction Stop"
            Write-LogEntry -logEntry "Remove-AppxProvisionedPackage -Online -PackageName '$($removeAppxProvisionedPackageName)' -AllUsers -ErrorAction Stop" -logID $logID 

            Remove-AppxProvisionedPackage -Online -PackageName $removeAppxProvisionedPackageName -AllUsers -ErrorAction Stop
        }
        catch [System.Exception] {
            Write-Warning -message "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed"
            Write-Warning -message $_.Exception.Message
            Write-LogEntry -logEntry "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed" -logID $logID -severity 3
            Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3

            $LASTEXITCODE = 1
        } 
    }

    function Test-WinGetBinary {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$winGetPackageName,
            [string]$winGetBinary
        )

        # Test the WinGet package and other dependancies are installed
        Write-Host "Testing the WinGet package and other dependancies are installed"
        Write-LogEntry -logEntry "Testing the WinGet package and other dependancies are installed" -logID $logID

        try {
            $winGetPath = (Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq $winGetPackageName }).InstallLocation | Sort-Object -Descending | Select-Object -First 1 -ErrorAction Stop
        }
        catch {
            $testWinGetPathFail = $true
            Write-Warning -Message "There was a problem getting details of the '$($winGetPackageName)' package"
            Write-Warning -message $_.Exception.Message
            Write-LogEntry -logEntry "There was a problem getting details of the '$($winGetPackageName)' package" -logID $logID -severity 3
            Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3
        }

        if ([string]::IsNullOrEmpty($winGetPath)) {
            $testWinGetPathFail = $true
            Write-Warning "The '$($winGetPackageName)' package was not found'"
            Write-LogEntry -logEntry "The '$($winGetPackageName)' package was not found" -logID $logID  -severity 3
        }
        else {

            $winGetBinaryPath = Join-Path -Path $winGetPath -ChildPath $winGetBinary

            try {
                if (Test-Path -Path $winGetBinaryPath ) {
                    Write-Host "The '$($winGetBinary)' binary was found at '$($winGetBinaryPath)'"
                    Write-LogEntry -logEntry "The '$($winGetBinary)' package was found at '$($winGetBinaryPath)'" -logID $logID
                }
                else {
                    $testWinGetBinaryPathFail = $true
                    Write-Warning "The '$($winGetPackageName)' package was found at '$($winGetPath)' but the WinGet binary was not found at '$($winGetBinaryPath)'"
                    Write-LogEntry -logEntry "The '$($winGetPackageName)' package was found at '$($winGetPath)' but the WinGet binary was not found at '$($winGetBinaryPath)'" -logID $logID -severity 3
                }
            }
            catch {
                $testWinGetBinaryPathFail = $true
                Write-Warning "An error was encounted trying to validate the path to WinGet.exe at '$($winGetBinaryPath)'"
                Write-Warning -message $_.Exception.Message
                Write-LogEntry -logEntry "An error was encounted trying to validate the path to WinGet.exe at '$($winGetBinaryPath)'" -logID $logID -severity 3
                Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3
            }
        }
        
        if ($testWinGetBinaryPathFail -or $testWinGetPathFail) {
            Write-Warning "The '$($winGetPackageName)' package was not found or the WinGet binary was not found at '$($winGetBinaryPath)'. Cannot continue"
            Write-LogEntry -logEntry "The '$($winGetPackageName)' package was not found or the WinGet binary was not found at '$($winGetBinaryPath)'. Cannot continue" -logID $logID -severity 3

            return @{Result = 'Fatal Error'; winGetPath = $null }
            $LASTEXITCODE = 1
        }
        else {
            
            # Test WinGet will run as SYSTEM
            try {
                Write-Host "Testing the WinGet binary will run as SYSTEM..."
                Write-LogEntry -logEntry "Testing the WinGet binary will run as SYSTEM..." -logID $logID
                Write-Host "$($winGetBinary) --version"
                Write-LogEntry -logEntry "$($winGetBinary) --version" -logID $logID

                Set-Location $winGetPath
                $testWinGetAsSystem = & .\$winGetBinary --version

                if (-not [string]::IsNullOrEmpty($testWinGetAsSystem)) {
                    Write-Host "The WinGet binary was validated"
                    Write-Host "WinGet version is '$($testWinGetAsSystem)'"
                    Write-LogEntry -logEntry "The WinGet binary was validated" -logID $logID
                    Write-LogEntry -logEntry "WinGet version is '$($testWinGetAsSystem)'" -logID $logID
                }
                else {
                    $testWinGetAsSystemFail = $true
                    Write-Warning "No output was detected while testing the WinGet version"
                    Write-Warning "WinGet has a dependencies, including the VC++ redistributable 14.x when running in the SYSTEM context. Ensure VC++ redistributable 14.x or higher and other dependencies are installed"
                    Write-LogEntry -logEntry "No output was detected while testing the WinGet version"-logID $logID -severity 3
                    Write-LogEntry -logEntry "WinGet has a dependencies, including the VC++ redistributable 14.x when running in the SYSTEM context. Ensure VC++ redistributable 14.x or higher and other dependencies are installed" -logID $logID -severity 3
                }
            }
            catch {
                $testWinGetAsSystemFail = $true
                Write-Warning "An error was encountered trying to run the WinGet binary using the command '$($testWinGetAsSystem)'"
                Write-Warning -message $_.Exception.Message
                Write-LogEntry -logEntry "An error was encountered trying to run the WinGet binary using the command '$($testWinGetAsSystem)'" -logID $logID -severity 3
                Write-LogEntry -logEntry $_.Exception.Message -logID $logID -severity 3
            }
        }

        if ($testWinGetAsSystemFail) {
            return @{Result = 'Failed'; winGetPath = $null }
        }
        else {
            return @{Result = 'Passed'; winGetPath = $winGetPath }
        }
        Set-Location $PSScriptRoot
    }

    function Test-WinGetApp {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$winGetAppName,
            [string]$winGetAppId,
            [string]$winGetAppSource,
            [string]$winGetPath,
            [string]$winGetBinary
        )
        try {
            Write-Host "Checking if '$($winGetAppName)' is installed using Id '$($winGetAppId)'..."
            Write-Host "winget.exe list --id $($winGetAppId) --source $($winGetAppSource) --accept-source-agreements"
            Write-LogEntry -logEntry "Checking if '$($winGetAppName)' is installed using Id '$($winGetAppId)'..." -logID $logID
            Write-LogEntry -logEntry "winget.exe list --id '$($winGetAppId)' --source $($winGetAppSource) --accept-source-agreements" -logID $logID 

            Set-Location $winGetPath
            $winGetTest = & .\$winGetBinary list --id $winGetAppId --source $winGetAppSource --accept-source-agreements
                
            foreach ($line in $winGetTest) {
                
                if ($line -like "*No installed package found*") {
                    Write-Host "The 'WinGet list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetAppId)', is not installed"
                    Write-LogEntry -logEntry "The 'WinGet list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetAppId)', is not installed" -logID $logID
                    
                    return @{Result = 'Not Installed' }
                }

                if ($line -like "*$($winGetAppId)*") {
                    Write-Host "The 'WinGet list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetAppId)', is installed"
                    Write-LogEntry -logEntry "The 'WinGet list' command line indicated the '$($winGetAppName)' app, with Id '$($winGetAppId)', is installed" -logID $logID
                    
                    return @{Result = 'Installed' }
                }
            }
        }
        catch {
            Write-Warning -Message "Error while running the WinGet command line to check if '$($winGetAppName)' is already installed"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "Error while running the WinGet command line to check if $($winGetAppName) is already installed" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3
            
            return @{Result = 'Fatal Error' }
            $LASTEXITCODE = 1
        }
        Set-Location $PSScriptRoot
    }
    function Install-WinGetApp {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false, ValuefromPipeline = $false)]
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [Parameter(Mandatory = $true, ValuefromPipeline = $true)]
            [string]$winGetAppName,
            [string]$winGetAppId,
            [string]$winGetAppSource,
            [string]$winGetPath,
            [string]$winGetBinary
        )
    
        # Attempt to install app using WinGet
        try {
            Write-Host "Installing '$($winGetAppName)' using the WinGet command line"
            Write-LogEntry -logEntry "Installing '$($winGetAppName)' using the WinGet command line" -logID $logID
            Write-Host ".\winget.exe install --name '$winGetAppName' --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine"
            Write-LogEntry -logEntry ".\winget.exe install --name '$winGetAppName' --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine" -logID $logID

            Set-Location $winGetPath
            & .\$winGetBinary install --name $winGetAppName --accept-package-agreements --accept-source-agreements --source $winGetAppSource --scope machine
        }
        catch {
            Write-Warning -Message "There was an error installing '$($winGetAppName)', with Id '$($winGetAppId)', using the WinGet command line"
            Write-Warning -Message "$($_.Exception.Message)"
            Write-LogEntry -logEntry "There was an error installing '$($winGetAppName)', with Id '$($winGetAppId)', using the WinGet command line" -logID $logID -severity 3
            Write-LogEntry -logEntry "$($_.Exception.Message)" -logID $logID -severity 3

            $LASTEXITCODE = 1
        }
        Set-Location $PSScriptRoot
    }

    # Initial logging
    Write-Host '** Starting processing the script' 
    Write-LogEntry -logEntry '** Starting processing the script' -logID $logID 
 
    # Call Functions  
    $testAppxPackage = Test-AppxPackage -removeApp $removeApp -appxWaitTimerSeconds $appxWaitTimerSeconds
    $testAppxProvisionedPackage = Test-AppxProvisionedPackage -removeApp $removeApp

    # Remove AppxPackage
    if ($testAppxPackage.Result -eq 'Installed') {
        $removeAppxPackage = Remove-AppxPkg -removeApp $removeApp

        # Check if AppxPackage needs registering before it can be removed
        if ($removeAppxPackage.Result -eq 'Failed' -and $removeAppxPackage.Reason -eq '0x80073CF1') {
            Register-AppxPackage -removeApp $removeApp -appxWaitTimerSeconds $appxWaitTimerSeconds
        }

        # Remove AppxProvionedPackage
        if ($testAppxProvisionedPackage.Result -eq 'Installed') {
            Remove-AppxProvPackage -removeApp $removeApp -removeAppxProvisionedPackageName $testAppxProvisionedPackage.PackageName
        }
    }
    
    # WinGet App intent is to be installed
    if ($winGetAppInstall -eq $true) {
        $winGetBinaryTestResult = Test-WinGetBinary -winGetBinary $winGetBinary -winGetPackageName $winGetPackageName
       
        if ($winGetBinaryTestResult.Result -eq 'Installed') {
            $winGetAppTest = Test-WinGetApp -winGetAppName $winGetAppName -winGetAppId $winGetAppId -winGetAppSource $winGetAppSource -winGetPath $winGetBinaryTestResult.winGetPath -winGetBinary $winGetBinary
           
            if ($winGetAppTest.Result -eq 'Not Installed') {

                # WinGet App is not installed. Install WinGet App
                Install-WinGetApp -winGetAppName $winGetAppName -winGetAppId $winGetAppId -winGetAppSource $winGetAppSource -winGetPath $winGetBinaryTestResult.winGetPath -winGetBinary $winGetBinary
            }
        } 

        # Test if WinGet App install retry is required
        $testAppxPackage = $null
        $winGetBinaryTestResult = $null
        $testAppxPackage = Test-AppxPackage -removeApp $removeApp -appxWaitTimerSeconds $appxWaitTimerSeconds
        $winGetBinaryTestResult = Test-WinGetBinary -winGetBinary $winGetBinary -winGetPackageName $winGetPackageName

        If ($winGetBinaryTestResult.Result -eq 'Not Installed' -or `
                $testAppxPackage.Result -eq 'SYSTEM Staged' -or $testAppxPackage.Result -eq 'Not Installed') {

            Write-Warning -Message "The WinGet app '$($winGetAppName)', is not installed correctly. Retrying the WinGet App install..."
            Write-LogEntry -logEntry "The WinGet app '$($winGetAppName)', is not installed correctly. Retrying the WinGet App install..." -logID $logID -severity 2

            # Retry WinGet App install
            $i = 1

            do {
                Write-Host "Retry attempt $i of $winGetRetries"
                Write-LogEntry -logEntry "Retry attempt $i of $winGetRetries" -logID $logID -severity 2
    
                # Test if WinGet App is installed
                $winGetAppTest = Test-WinGetApp -winGetAppName $winGetAppName -winGetAppId $winGetAppId -winGetAppSource $winGetAppSource -winGetPath $winGetBinaryTestResult.winGetPath -winGetBinary $winGetBinary
    
                # WinGet App is not installed. Install WinGet App
                if ($winGetAppTest.Result -eq 'Not Installed') {
                    $i++
                    Install-WinGetApp -winGetAppName $winGetAppName -winGetAppId $winGetAppId -winGetAppSource $winGetAppSource -winGetPath $winGetBinaryTestResult.winGetPath -winGetBinary $winGetBinary
                    $winGetAppTest = $null
                    $winGetAppTest = Test-WinGetApp -winGetAppName $winGetAppName -winGetAppId $winGetAppId -winGetAppSource $winGetAppSource -winGetPath $winGetBinaryTestResult.winGetPath -winGetBinary $winGetBinary
                }
            }

            while ($i -le $winGetRetries -and (-not $winGetAppTest.Result -eq 'Installed'))

            If ($i -eq $winGetRetries -and (-not $winGetAppTest -eq 'Installed')) {
                Write-Warning -Message "The WinGet app '$($winGetAppName)', did not install correctly after '$($winGetRetries) attempts. The maximum number of retries has been reached"
                Write-LogEntry -logEntry "The WinGet app '$($winGetAppName)', did not install correctly after '$($winGetRetries) attempts. The maximum number of retries has been reached" -logID $logID -severity 3
            
                $LASTEXITCODE = 1

            }
            elseif ($winGetAppTest -eq 'Installed') {  
                Write-Host "The WinGet app '$($winGetAppName)', installed correctly after '$($i) attempts"
                Write-LogEntry -logEntry "The WinGet app '$($winGetAppName)', installed correctly after '$($i) attempts" -logID $logID -severity 1
            }
        }
    }
}
end {

    # Complete
    Write-Output "Finished processing the script"
    Write-LogEntry -logEntry "Finished processing the script" -logID $logID
    Set-Location $PSScriptRoot
}