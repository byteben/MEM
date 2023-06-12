<#
.SYNOPSIS
    Remove a built-in modern app from Windows, for All Users, and reinstall using WinGet
    
.DESCRIPTION
    This script will remove a specific built-in AppXPackage, for All Users, and also the AppXProvisionedPackage
    When deploying apps from the new store, via Intune, in the SYSTEM context, an error appears for the deployment if the app was previous deployed in the USER context
    "The application was not detect after installation complete successfully (0x87D1041C)"
    This script will remove all existing instances of the app so the app from Intune can be installed sucessfully
    AppxPackage removal can fail if the app was installed from the Microsoft Store. This script will re-register the app for All Users in that instance to allow for removal

    .NOTES
    FileName:   Remove-Appx.ps1
    Date:       12th June 2023
    Author:     Ben Whitmore @ PatchMyPC (Thanks to Bryan Dam @bdam555 for assisted research)
    Contact:    @byteben
    
.PARAMETER removeApp
    Specify the AppxPackage and AppxProivisionedPackage to remove
    This is not required. The parameter is defined at the top of the script so it can be used as an Intune Script (which does not accept params)

.PARAMETER reinstallApp
    Specify app to reinstall using WinGet
    This is not required. The parameter is defined at the top of the script so it can be used as an Intune Script (which does not accept params)

.PARAMETER reinstallSource
    Specify WinGet source to use. Typically this will be msstore
    This is not required. The parameter is defined at the top of the script so it can be used as an Intune Script (which does not accept params)

.EXAMPLE
    .\Reset-Appx.ps1

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$removeApp = 'Microsoft.CompanyPortal',
    [string]$reinstallApp = 'Company Portal',
    [string]$reinstallAppSource = 'msstore',
    [string]$logID = 'Main'
)

Begin {

    #Create variables
    $removeAppxPackage = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp }
    $removeAppxProvisionedPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $removeApp } | Select-Object -ExpandProperty PackageName
    $removeAppxPackageError = $null

    #Time format
    $dateTime = Get-Date
    $date = $dateTime.ToString("MM-dd-yyyy", [Globalization.CultureInfo]::InvariantCulture)
    $time = $dateTime.ToString("HH:mm:ss", [Globalization.CultureInfo]::InvariantCulture)
}

Process {

    # Functions
    function Write-LogEntry {
        param(
            [parameter(Mandatory = $true)]
            [string[]]$logEntry,
            [string]$logID,
            [parameter(Mandatory = $false)]
            [string]$logFile = "$($env:temp)\Reset-Appx.log",
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
                #Extract log object and construct format for log line entry
                foreach ($log in $logEntry) {
                    $logDetail = [string]::Format('<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">', $log, $time, $date, $component, $context, $severity, $PID)

                    #Attempt log write
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
        If (-not[string]::IsNullOrEmpty($removeAppxPackage)) {
           
            try {
                
                #List users with the AppxPackage installed
                Write-Host "Found the following users with AppxPackage '$($removeApp)' installed:"
                Write-LogEntry -logEntry "Found the following users with AppxPackage '$($removeApp)' installed:" -logID $logID 
                
                foreach ($removeAppxPackageUserInfo in $removeAppxPackage.PackageUserInformation) {
                    $removeAppxPackageUser = ( $removeAppxPackageUserInfo | Where-Object { $_.InstallState -like 'Installed*' } | Select-Object -ExpandProperty UserSecurityId).UserName
                    Write-Host $removeAppxPackageUser
                    Write-LogEntry -logEntry $removeAppxPackageUser -logID $logID 
                }

                Write-Host "Removing AppxPackage: $($removeApp)"
                Write-LogEntry -logEntry "Removing AppxPackage: $($removeApp)" -logID $logID  
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            catch [System.Exception] {
                Write-Warning "AppxPackage removal failed. AppxPackage '$($removeApp)' needs to be re-registered before it can be removed."
                Write-LogEntry -logEntry "AppxPackage removal failed. AppxPackage '$($removeApp)' needs to be re-registered before it can be removed." -logID $logID 
                
                if ( $_.Exception.Message -like "*HRESULT: 0x80073CF1*") {
                    $removeAppxPackageError = $true
                }
                else {
                    Write-Warning -Message "Removing AppxPackage '$($removeApp)' failed: $($_.Exception.Message)"
                    Write-LogEntry -logEntry "Removing AppxPackage '$($removeApp)' failed: $($_.Exception.Message)" -logID $logID -severity 3
                }
            }
            
            #Test removal was successful
            $testAppx = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp }

            If ([string]::IsNullOrEmpty($testAppx)) {
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

        #Re-register AppxPackage for all users and attempt removal again
        if ( $removeAppxPackageError ) {
            Write-Host "Attempting to re-register AppxPackage '$($removeApp)'..."
            Write-LogEntry -logEntry "Attempting to re-register AppxPackage '$($removeApp)'..." -logID $logID 

            Try {
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }
                If (-not[string]::IsNullOrEmpty( { Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } } )) {
                    Write-Host "AppxPackage '$($removeApp)' registered succesfully"
                    Write-LogEntry -logEntry "AppxPackage '$($removeApp)' registered succesfully" -logID $logID 
                }
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $removeApp } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            Catch {
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
                Write-LogEntry -logEntry "Removing AppxProvisioningPackage: $($removeAppxProvisionedPackageName)" -logID $logID 
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $removeAppxProvisioningPackageName } | Remove-AppxProvisionedPackage -AllUsers -ErrorAction Stop | Out-Null
            }
            catch [System.Exception] {
                Write-Warning -message "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed: $($_.Exception.Message)"
                Write-LogEntry -logEntry "Removing AppxProvisionedPackage '$($removeAppxProvisionedPackageName)' failed: $($_.Exception.Message)" -logID $logID -severity 3
            }
        }
        else {
            Write-Output "Did not attempt removal of the AppxProvisionedPackage '$($removeApp)' because it was not found"
            Write-LogEntry -logEntry "Did not attempt removal of the AppxProvisionedPackage '$($removeApp)' because it was not found" -logID $logID -severity 2
        }

        #Test removal was successful
        $testAppxProv = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $removeAppxProvisioningPackageName }

        If ([string]::IsNullOrEmpty($testAppxProv)) {
            Write-Host "AppxProvisionedPackage: $($removeApp) was removed succesfully"
            Write-LogEntry -logEntry "AppxProvisionedPackage: $($removeApp) was removed succesfully" -logID $logID  
        }
        else {
            Write-Warning -Message "AppxProvisionedPackage: $($removeApp) removal was unsuccessful"
            Write-LogEntry -logEntry "AppxProvisionedPackage: $($removeApp) removal was unsuccessful" -logID $logID -severity 3
        }
    }

    function Install-WinGetApp {
        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name,
            [string]$winGetPath = 'C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe'
        )
    
        # Attempt to install app using WinGet
        If (Test-Path -Path $winGetPath) {

            Write-Host "Installing $($reinstallApp) using WinGet command line..."
            Write-LogEntry -logEntry "Installing $($reinstallApp) using WinGet command line..." -logID $logID 

            Set-Location (Resolve-Path $winGetPath).path
            .\winget.exe install --Name $reinstallApp --accept-package-agreements --accept-source-agreements --exact --source $reinstallAppSource --scope machine

            if (-not[string]::IsNullOrEmpty({ .\winget.exe list $reinstallApp --source $reinstallAppSource })) {
                Write-Host "Sucessfully installed $($reinstallApp) using WinGet command line"
                Write-LogEntry -logEntry "Sucessfully installed $($reinstallApp) using WinGet command line" -logID $logID 
            }
            else {
                Write-Host "$($reinstallApp) not found after attempting install with WinGet command line..."
                Write-LogEntry -logEntry "$($reinstallApp) not found after attempting install with WinGet command line..." -logID $logID -severity 2
            }
        }
        else {
            Write-Warning 'WinGet.exe not found. Aborting reinstall attempt'
            Write-LogEntry -logEntry 'WinGet.exe not found. Aborting reinstall attempt' -logID $logID -severity 3
        }
    }

    # Initial logging
    Write-Host 'Starting AppxPackage and AppxProvisionedPackage removal process'
    Write-Host "Processing AppxPackage: $($removeApp)"
    Write-LogEntry -logEntry "** Starting AppxPackage and AppxProvisionedPackage removal process" -logID $logID 
    Write-LogEntry -logEntry "Processing appx package: $($removeApp)" -logID $logID  

    # Call Functions
    Remove-AppxPkg
    Remove-AppxProvPkg

    if ($reinstallApp -and $reinstallAppSource) {
        Install-WinGetApp
    }

    # Complete
    Write-Output "Completed built-in AppxPackage and AppxProvisionedPackage removal process"
    Write-LogEntry -logEntry "Completed built-in AppxPackage and AppxProvisionedPackage removal process" -logID $logID
    Set-Location $PSScriptRoot
}