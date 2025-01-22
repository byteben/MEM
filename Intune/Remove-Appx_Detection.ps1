
Begin {

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

    $AppArrayList = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName

    foreach ($App in $AppArrayList) {
        if (($App -in $BlackListedApps)) {
            $AppExists = $True
        }
    }

    If ($AppExists) {
        Write-Output "All appx packages were not removed"
        Exit 1
    }
    else {
        Write-Output "All appx packages were removed"
        Exit 0
    }
}