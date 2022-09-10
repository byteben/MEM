<#
.SYNOPSIS
    Detection Script for unsupported applications
.DESCRIPTION
    This script is designed to be run as a Proactive Remediation. 
    The BadApps array contains application names that are considered unsupported by the company.
    If no BadApps are found, the exit code of the script is 0 (Do not remediate). If BadApps are found, the exit code is 1 (Remediate)
.EXAMPLE
    Detect-JavaApps.ps1 (Run in the User Context)      
.NOTES
    FileName:    Detect-JavaApps.ps1
    Author:      Ben Whitmore
    Contributor: Jan Ketil Skanke
    Contact:     @byteben
    Created:     2022-11-Mar

    Version history:
    2.0.0 - (2022-03-01) Split Detection/Remeiation scripts and used Invoke-WebRequest instead of WebClient .NET http download
#>

#region SCRIPTVARIABLES
$BadApps = @(
    "Java 6"
    "Java SE Development Kit 6"
    "Java(TM) SE Development Kit 6"
    "Java(TM) 6"
    "Java 7"
    "Java SE Development Kit 17"
    "Java(TM) SE Development Kit 17"
    "Java(TM) 7"
)
#endregion

# Function to get all Installed Applications
function Get-InstalledApplications() {
    param(
        [string]$UserSid
    )
    
    New-PSDrive -PSProvider Registry -Name "HKU" -Root HKEY_USERS | Out-Null
    $regpath = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $regpath += "HKU:\$UserSid\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    if (-not ([IntPtr]::Size -eq 4)) {
        $regpath += "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $regpath += "HKU:\$UserSid\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    }
    $PropertyNames = 'DisplayName', 'DisplayVersion', 'Publisher', 'UninstallString'
    $Apps = Get-ItemProperty $regpath -Name $PropertyNames -ErrorAction SilentlyContinue | . { process { if ($_.DisplayName) { $_ } } } | Select-Object DisplayName, DisplayVersion, Publisher | Sort-Object DisplayName   
    Remove-PSDrive -Name "HKU" | Out-Null
    Return $Apps
}
#end function

#region GETSID
#Get SID of current interactive users
$CurrentLoggedOnUser = (Get-CimInstance win32_computersystem).UserName
if (-not ([string]::IsNullOrEmpty($CurrentLoggedOnUser))) {
    $AdObj = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
    $strSID = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
    $UserSid = $strSID.Value
}
else {
    $UserSid = $null
}
#endregion
	
#region APPINVENTORY
#Get Apps for system and current user
$MyApps = Get-InstalledApplications -UserSid $UserSid
$UniqueApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -eq 1 }).Group
$DuplicatedApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -gt 1 }).Group
$NewestDuplicateApp = ($DuplicatedApps | Group-Object DisplayName) | ForEach-Object { $_.Group | Sort-Object [version]DisplayVersion -Descending | Select-Object -First 1 }
$CleanAppList = $UniqueApps + $NewestDuplicateApp | Sort-Object DisplayName

#Build App Array
$AppArray = @()
foreach ($App in $CleanAppList) {
    $tempapp = New-Object -TypeName PSObject
    $tempapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $App.DisplayName -Force
    $tempapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value $App.DisplayVersion -Force
    $tempapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value $App.Publisher -Force
    $AppArray += $tempapp
}	
$AppPayLoad = $AppArray
#endregion APPINVENTORY

#region Find Bad Apps
$BadAppFound = $Null
$BadAppArray = @()

Foreach ($App in $AppPayLoad) {
    Foreach ($BadApp in $BadApps) {
        If ($App.AppName -like "*$BadApp*") {
            $tempbadapp = New-Object -TypeName PSObject
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $App.AppName -Force
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value $App.AppVersion -Force
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value $App.AppPublisher -Force
            $BadAppArray += $tempbadapp
            $BadAppFound = $True
        }
    }
}
$BadAppPayLoad = $BadAppArray

If ($BadAppFound) {
    #Write-Output for Proactive Remediation
    $BadAppPayLoadOutput = $BadAppPayLoad | ConvertTo-Json -Compress
    Write-Output $BadAppPayLoadOutput
    Exit 1
}
else {
    Write-Output "OK"
    Exit 0
}
#endregion