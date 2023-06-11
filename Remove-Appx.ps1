<#
.SYNOPSIS
    Remove a built-in modern app from Windows, for All Users. 
    
.DESCRIPTION
    This script will remove a speciifc built-in AppXPackage, for All Users, and also the AppXProvisionedPackage
    When deploying apps from the new store, via Intune, in the SYSTEM context, an error appears for the deployment if the app was previous deployed in the USER context
    "The application was not detect after installation complete successfully (0x87D1041C)"
    This script will remove all existing instances of the app so the app from Intune can be installed sucessfully
    AppxPackage removal can fail if the app was installed from the Microsoft Store. This script will re-register the app for All Users in that instance to allow for removal

    .NOTES
    FileName:    Remove-Appx.ps1
    Author:      Ben Whitmore @ PatchMyPC (Thanks to Bryan Dam @bdam555 for assisted research)
    Contact:     @byteben
    
.PARAMETER app
    Specify the AppxPackage and AppxProivisionedPackage to remove
    This is not required. The parameter is defined at the top of the script so it can be used as an Intune Script (which does not accept params)

.EXAMPLE
    .\Remove-Appx.ps1

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$app = 'Microsoft.CompanyPortal',
    [string]$logID = 'Main'
)

Begin {

    #Create variables
    $appxPackage = Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $app }
    $appxProvisionedPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $app } | Select-Object -ExpandProperty PackageName
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
            [string]$logFile = "$($env:temp)\Remove-Appx.log",
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
            # Add value to log file
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
        If (-not[string]::IsNullOrEmpty($appxPackage)) {
            try {
                
                #List users with the AppxPackage installed
                Write-Host "Found the following users with AppxPackage '$($app)' installed:"
                Write-LogEntry -logEntry "Found the following users with AppxPackage '$($app)' installed:" -logID $logID 
                
                foreach ($appxPackageUserInfo in $appxPackage.PackageUserInformation) {

                    $appxPackageUser = ( $appxPackageUserInfo | Where-Object { $_.InstallState -like 'Installed*' } | Select-Object -ExpandProperty UserSecurityId).UserName
                    Write-Host $appxPackageUser
                    Write-LogEntry -logEntry $appxPackageUser -logID $logID 
                }

                Write-Host "Removing AppxPackage: $($app)"
                Write-LogEntry -logEntry "Removing AppxPackage: $($app)" -logID $logID  
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $app } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            catch [System.Exception] {
                Write-Warning "AppxPackage removal failed. AppxPackage '$($app)' needs to be re-registered before it can be removed."
                Write-LogEntry -logEntry "AppxPackage removal failed. AppxPackage '$($app)' needs to be re-registered before it can be removed." -logID $logID 
                if ( $_.Exception.Message -like "*HRESULT: 0x80073CF1*") {
                    $removeAppxPackageError = $true
                }
                else {
                    Write-Warning -Message "Removing AppxPackage '$($app)' failed: $($_.Exception.Message)"
                    Write-LogEntry -logEntry "Removing AppxPackage '$($app)' failed: $($_.Exception.Message)" -logID $logID -severity 3
        
                }
            }
        }
        else {
            Write-Output "Did not attempt removal of the AppxPackage '$($app)' because it was not found"
            Write-LogEntry -logEntry "Did not attempt removal of the AppxPackage '$($app)' because it was not found" -logID $logID -severity 2
        }

        #Re-register AppxPackage for all users and attempt removal again
        if ( $removeAppxPackageError ) {
            Write-Host "Attempting to re-register AppxPackage '$($app)'..."
            Write-LogEntry -logEntry "Attempting to re-register AppxPackage '$($app)'..." -logID $logID 

            Try {
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $app } | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }
                If (-not[string]::IsNullOrEmpty( { Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $app } } )) {
                    Write-Host "AppxPackage '$($app)' registered succesfully"
                    Write-LogEntry -logEntry "AppxPackage '$($app)' registered succesfully" -logID $logID 
                }
                Get-AppXPackage -AllUsers | Where-Object { $_.Name -like $app } | Remove-AppxPackage -AllUsers -ErrorAction Stop
            }
            Catch {
                Write-Warning -Message "Re-registering AppxPackage '$($app)' failed: $($_.Exception.Message)"
                Write-LogEntry -logEntry "Re-registering AppxPackage '$($app)' failed: $($_.Exception.Message)" -logID $logID -severity 3
            }
        }
    }

    function Remove-AppxProvPkg {
        [CmdletBinding()]
        Param(
            [string]$logID = $($MyInvocation.MyCommand).Name
        )
    
        # Attempt to remove AppxProvisionedPackage
        If (-not[string]::IsNullOrEmpty($appxProvisionedPackageName)) {

            try {
                Write-LogEntry -logEntry "Removing AppxProvisioningPackage: $($appxProvisionedPackageName)" -logID $logID 
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $appxProvisioningPackageName } | Remove-AppxProvisionedPackage -AllUsers -ErrorAction Stop | Out-Null
            }
            catch [System.Exception] {
                Write-Warning -message "Removing AppxProvisionedPackage '$($appxProvisionedPackageName)' failed: $($_.Exception.Message)"
                Write-LogEntry -logEntry "Removing AppxProvisionedPackage '$($appxProvisionedPackageName)' failed: $($_.Exception.Message)" -logID $logID -severity 3
            }
        }
        else {
            Write-Output "Did not attempt removal of the AppxProvisionedPackage '$($app)' because it was not found"
            Write-LogEntry -logEntry "Did not attempt removal of the AppxProvisionedPackage '$($app)' because it was not found" -logID $logID -severity 2
        }
    }

    # Initial logging
    Write-Host 'Starting AppxPackage and AppxProvisionedPackage removal process'
    Write-Host "Processing AppxPackage: $($app)"
    Write-LogEntry -logEntry '##################' -logID $logID 
    Write-LogEntry -logEntry 'Starting AppxPackage and AppxProvisionedPackage removal process' -logID $logID 
    Write-LogEntry -logEntry "### $($date) $($time)  ###" -logID $logID 
    Write-LogEntry -logEntry '##################' -logID $logID 
    Write-LogEntry -logEntry "Processing appx package: $($app)" -logID $logID  

    # Call Functions
    Remove-AppxPkg
    Remove-AppxProvPkg

    # Complete
    Write-Output "Completed built-in AppxPackage and AppxProvisionedPackage removal process"
    Write-LogEntry -logEntry "Completed built-in AppxPackage and AppxProvisionedPackage removal process" -logID $logID
}