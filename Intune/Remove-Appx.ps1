<#
.SYNOPSIS
    Remove built-in apps (modern apps) from Windows 10.
.DESCRIPTION
    This script will remove all built-in apps with a provisioning package that are specified in the 'black-list' in this script.
.EXAMPLE
    .\Remove-Appx.ps1
.NOTES

    Based on original script / Credit to:-
    FileName:    Invoke-RemoveBuiltinApps.ps1
    Author:      Nickolaj Andersen

    Modifications to original script to Black list Appx instead of Whitelist

    FileName:    Remove-Apps.ps1
    Author:      Ben Whitmore
    Contact:     @byteben

    ################################
    Appx Packages as of 27/08/2021
    ################################

    Full Package List: https://docs.microsoft.com/en-us/windows/application-management/apps-in-windows-10

    ## 1809 - 21H1 ##
    Microsoft.BingWeather
    Microsoft.DesktopAppInstaller
    Microsoft.GetHelp
    Microsoft.GetStarted
    Microsoft.HEIFImageExtension
    Microsoft.Microsoft3DViewer
    Microsoft.MicrosoftOfficeHub
    Microsoft.MicrosoftSolitaireCollection
    Microsoft.MicrosoftStickyNotes
    Microsoft.MixedReality.Portal
    Microsoft.MSPaint
    Microsoft.Office.OneNote
    Microsoft.People
    Microsoft.ScreenSketch
    Microsoft.SkypeApp
    Microsoft.StorePurchaseApp
    Microsoft.VP9VideoExtensions
    Microsoft.Wallet
    Microsoft.WebMediaExtensions
    Microsoft.Windows.Photos
    Microsoft.WindowsAlarms
    Microsoft.WindowsCalculator
    Microsoft.WindowsCamera
    Microsoft.WindowsCommunicationsApps
    Microsoft.WindowsFeedbackHub
    Microsoft.WindowsMaps
    Microsoft.WindowsSoundRecorder
    Microsoft.WindowsStore
    Microsoft.Xbox.TCUI
    Microsoft.XboxApp
    Microsoft.XboxGameOverlay
    Microsoft.XboxGamingOverlay
    Microsoft.XboxIdentityProvider
    Microsoft.XboxSpeechToTextOverlay
    Microsoft.YourPhone
    Microsoft.ZuneMusic
    Microsoft.ZuneVideo

    ## 1809, 1903, 1909, 20H2, 21H1 ##
    Microsoft.Messaging
    Microsoft.OneConnect
    Microsoft.Print3D

    ## 1909, 20H2, 21H1##
    Microsoft.Outlook.DesktopIntegrationServices

    ## 21H1 ##
    Microsoft.3DBuilder
#>
Begin {

    # Black list of appx packages to remove
    $BlackListedApps = New-Object -TypeName System.Collections.ArrayList
    $BlackListedApps.AddRange(@(
        "Microsoft.GetHelp",
        "Microsoft.GetStarted",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.MixedReality.Portal",
        "Microsoft.SkypeApp",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.XboxApp",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo"
    ))
}
Process {
    # Functions
    function Write-LogEntry {
        param(
            [parameter(Mandatory=$true, HelpMessage="Value added to the RemovedApps.log file.")]
            [ValidateNotNullOrEmpty()]
            [string]$Value,

            [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
            [ValidateNotNullOrEmpty()]
            [string]$FileName = "RemovedApps.log"
        )
        # Determine log file location
        $LogFilePath = Join-Path -Path $env:windir -ChildPath "Temp\$($FileName)"

        # Add value to log file
        try {
            Out-File -InputObject $Value -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        }
        catch [System.Exception] {
            Write-Warning -Message "Unable to append log entry to $($FileName) file"
        }
    }

    # Initial logging
    Write-LogEntry -Value "Starting built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process"

    # Determine provisioned apps
    $AppArrayList = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName

    # Loop through the list of appx packages
    foreach ($App in $AppArrayList) {
        Write-LogEntry -Value "Processing appx package: $($App)"

        # If application name not in appx package black list, remove AppxPackage and AppxProvisioningPackage
        if (($App -Notin $BlackListedApps)) {
            Write-LogEntry -Value "Skipping excluded application package: $($App)"
        }
        else {
            # Gather package names
            $AppPackageFullName = Get-AppxPackage -Name $App | Select-Object -ExpandProperty PackageFullName -First 1
            $AppProvisioningPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $App } | Select-Object -ExpandProperty PackageName -First 1

            # Attempt to remove AppxPackage
            if ($AppPackageFullName -ne $null) {
                try {
                    Write-LogEntry -Value "Removing AppxPackage: $($AppPackageFullName)"
                    Remove-AppxPackage -Package $AppPackageFullName -ErrorAction Stop | Out-Null
                }
                catch [System.Exception] {
                    Write-LogEntry -Value "Removing AppxPackage '$($AppPackageFullName)' failed: $($_.Exception.Message)"
                }
            }
            else {
                Write-LogEntry -Value "Unable to locate AppxPackage for current app: $($App)"
            }

            # Attempt to remove AppxProvisioningPackage
            if ($AppProvisioningPackageName -ne $null) {
                try {
                    Write-LogEntry -Value "Removing AppxProvisioningPackage: $($AppProvisioningPackageName)"
                    Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Online -ErrorAction Stop | Out-Null
                }
                catch [System.Exception] {
                    Write-LogEntry -Value "Removing AppxProvisioningPackage '$($AppProvisioningPackageName)' failed: $($_.Exception.Message)"
                }
            }
            else {
                Write-LogEntry -Value "Unable to locate AppxProvisioningPackage for current app: $($App)"
            }
        }
    }

    # Complete
    Write-LogEntry -Value "Completed built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process"
}