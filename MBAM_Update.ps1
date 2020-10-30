#Set Current Directory
$ScriptPath = $MyInvocation.MyCommand.Path
$CurrentDir = Split-Path $ScriptPath

#Set Variables
$MDOPUpdateRequired = $False
$MDOPMSI = Join-Path $CurrentDir "MbamClientSetup-2.5.1100.0.msi"
$MDOPMSP = Join-Path $CurrentDir "MBAM2.5_Client_x64_KB4586232.msp"
$MDOPVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4}" -ErrorAction SilentlyContinue).DisplayVersion 

If ($MDOPVersion) {

    Switch ($MDOPVersion) {
        "2.5.1100.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 Initial Package"
            $MDOPUpdateRequired = $True
        }

        "2.5.1126.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 December 2016 Package"
            $MDOPUpdateRequired = $True
        }

        "2.5.1133.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 March 2017 Package"
            $MDOPUpdateRequired = $True
        }

        "2.5.1134.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 June 2017 Package"
            $MDOPUpdateRequired = $True
        }

        "2.5.1143.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 July 2018 Package"
            $MDOPUpdateRequired = $True
        }

        "2.5.1147.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 March 2019 Package" 
            $MDOPUpdateRequired = $True
        }

        "2.5.1152.0" {
            Write-Output "MDOP is installed, the version is $MDOPVersion, 2.5SP1 October 2020 Package"
            $MDOPUpdateRequired = $False
        }
    
        Default {
            Write-Warning "MDOP Version could not be read"
            Exit 1
        }
    }
}
else {

    Write-Output "MDOP Doesnt Exist. Installing MDOP 2.5SP1 Base 2.5.1100.0...."

    #Install MDOP 2.5SP1 Base 2.5.1100.0
    $Args = @(
        "/i"
        """$MDOPMSI"""
        "/qn"
    )
    Start-Process msiexec -ArgumentList $Args -Wait -ErrorAction Stop | Out-Null
    $MDOPUpdateRequired = $True
}

If ($MDOPUpdateRequired -eq $True) {

    Write-Output "Installing MDOP 2.5SP1 October 2020 Package 2.5.1152.0...."

    #Install MDOP 2.5SP1 October 2020 Patch
    $Args = @(
        "/p"
        """$MDOPMSP"""
        "/qn"
    )
    Start-Process msiexec -ArgumentList $Args -Wait -ErrorAction Stop | Out-Null
}