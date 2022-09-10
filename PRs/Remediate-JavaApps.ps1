<#
.SYNOPSIS
    Remediation Script for unsupported applications
.DESCRIPTION
    This script is designed to be run as a Proactive Remediation. 
    The BadApps array contains application names that are considered unsupported by the company.
    If BadApps are found, a Toast Notification is displayed to the user informing them to remove the unsupported applications.
.EXAMPLE
    RemediateJavaApps.ps1 (Run in the User Context)      
.NOTES
    FileName:    Remediate-JavaApps.ps1
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

$CustomHandlerDisplayName = "Ninja Notifications"
$CustomHandlerAppID = "CustomToastNotify"
$GoodMorning = "Good Morning"
$GoodAfternoon = "Good Afternoon"
$GoodEvening = "Good Evening"
$ToastImageSource = "https://github.com/byteben/MEM/raw/main/PRs/ninja-banner.png" #ToastImage should be  364px x 180px
$ToastImage = Join-Path -Path $ENV:temp -ChildPath "ToastImage.jpg" #ToastImageSource is downloaded to this location
$ToastDuration = "long" #ToastDuration: Short = 7s, Long = 25s
$ToastScenario = "reminder" #ToastScenario: Default | Reminder | Alarm
$ToastTitle = "Java apps found on your computer"
$ToastText = "If not required, please uninstall the following Java applications at your earliest convenience as they pose a security risk to your computer:-"
$SnoozeTitle = "Set Reminder"
$SnoozeMessage = "Remind me again in"
$LogFile = Join-Path -Path $env:TEMP -ChildPath "UnsupportedAppsFound-Java.log"
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
# Function Write Log Entry
function Write-LogEntry {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = $($LogName),
        [switch]$Stamp
    )

    #Build Log File appending System Date/Time to output
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $Date = (Get-Date -Format "MM-dd-yyyy")

    If ($Stamp) {
        $LogText = "<$($Value)> <time=""$($Time)"" date=""$($Date)"">"
    }
    else {
        $LogText = "$($Value)"   
    }
	
    Try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFile -ErrorAction Stop
    }
    Catch [System.Exception] {
        Write-Warning -Message "Unable to add log entry to $LogFile.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}
#end function

#region RESETLOG
If (Test-Path -Path $LogFile) {
    Remove-Item $LogFile -Force | Out-Null
}
#endregion

#region FETCHIMAGE
New-Object uri $ToastImageSource

#Test if URL is valid
Try {
    #Attempt URL get and set Status Code variable
    $URLRequest = Invoke-WebRequest -UseBasicParsing -URI $ToastImageSource -ErrorAction SilentlyContinue
    $StatusCode = $URLRequest.StatusCode
    Write-LogEntry -Stamp -Value "Attempting to download $ToastImageSource"
    Write-LogEntry -Stamp -Value "Download status code is $StatusCode"
}
Catch {
    #Catch Status Code on error
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Exit 1
}

#If URL exists
If ($StatusCode -eq 200) {
    #Attempt File download
    Try {
        Invoke-WebRequest -UseBasicParsing -Uri $ToastImageSource -OutFile $ToastImage -ErrorAction SilentlyContinue
        #If download was successful, test the file was saved to the correct directory
        If (!(Test-Path -Path $ToastImage)) {
            Write-LogEntry -Stamp -Value "The download was interrupted or an error occured moving the file to the destination specified"
            Write-Output "The download was interrupted or an error occured moving the file to the destination specified"
            Exit 1
        }
    }
    Catch {
        #Catch any errors during the file download
        Write-LogEntry -Stamp -Value "Error downloading file: $ToastImageSource"
        write-Output "Error downloading file: $ToastImageSource" 
        Exit 1
    }
}
else {
    #For anything other than status 200 (URL OK), throw a warning
    Write-LogEntry -Stamp -Value "URL Does not exists or the website is down. Status Code: $StatusCode" 
    Write-Output "URL Does not exists or the website is down. Status Code: $StatusCode" 
    Exit 1
}
#endregion

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
$AppPayLoadLog = $AppPayLoad | Out-String
Write-LogEntry -Value "################Unique Apps Found################"
Write-LogEntry -Stamp -Value $AppPayLoadLog
#endregion APPINVENTORY

#region Find Bad Apps
$BadAppsLog = $BadApps | Out-String
Write-LogEntry -Value "################Unsupported Apps being searched for################"
Write-LogEntry -Stamp -Value $BadAppsLog

$BadAppArray = @()

Foreach ($App in $AppPayLoad) {
    Foreach ($BadApp in $BadApps) {
        If ($App.AppName -like "*$BadApp*") {
            $tempbadapp = New-Object -TypeName PSObject
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $App.AppName -Force
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value $App.AppVersion -Force
            $tempbadapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value $App.AppPublisher -Force
            $BadAppArray += $tempbadapp
        }
    }
}
$BadAppPayLoad = $BadAppArray

#Update Event Text Message to include bad apps
$EventText = $EventText + "`n"
Foreach ($BadApp2 in $BadAppPayload) { 
    $EventText = $EventText + "`n- $($BadApp2.AppName)"
}
Write-LogEntry -Value "################Toast Notification Details################"
Write-LogEntry -Stamp -Value $EventText
#endregion

$BadAppPayLoadLog = $BadAppPayLoad | Out-String
Write-LogEntry -Value "################Unsupported Apps Found################"
Write-LogEntry -Stamp -Value $BadAppPayLoadLog

#region CUSTOMHANDLER
#https://docs.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/send-local-toast-other-apps
$CustomToastNotifyRegKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\$CustomHandlerAppID"
$CustomHandlerClassRegKey = "HKCU:\Software\Classes\AppUserModelId"

Try {
    If (!(Test-Path -Path $CustomToastNotifyRegKey)) {
        New-Item -Path $CustomToastNotifyRegKey -Force | Out-Null
        New-ItemProperty -Path $CustomToastNotifyRegKey -Name "ShowInActionCenter" -Value 1 -PropertyType DWORD -Force | Out-Null
    }
}
Catch { 
    $_.Exception.Message 
}

Try {
    If (!(Test-Path -Path $CustomHandlerClassRegKey)) {
        New-Item -Path $CustomHandlerClassRegKey -Name $CustomHandlerAppID -Force | Out-Null
        New-ItemProperty -Path $CustomHandlerClassRegKey\$CustomHandlerAppID -Name "DisplayName" -Value $CustomHandlerDisplayName -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $CustomHandlerClassRegKey\$CustomHandlerAppID -Name "ShowInSettings" -Value 0 -PropertyType DWORD -Force | Out-Null
    }
}
Catch { 
    $_.Exception.Message 
}

Try {
    If ((Get-ItemProperty -Path $CustomHandlerClassRegKey\$CustomHandlerAppID -Name "DisplayName" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue) -ne $CustomHandlerDisplayName) {
        New-ItemProperty -Path $CustomHandlerClassRegKey\$CustomHandlerAppID -Name "DisplayName" -Value $CustomHandlerDisplayName -PropertyType String -Force | Out-Null
    }
}
Catch {
    $_.Exception.Message
}
#endregion

#region TOAST
#Get Hour of Day and set Custom Hello
$Hour = (Get-Date).Hour
If ($Hour -lt 12) { $CustomHello = $GoodMorning + ". " }
ElseIf ($Hour -gt 16) { $CustomHello = $GoodEvening + ". " }
Else { $CustomHello = $GoodAfternoon + ". " }

$CustomHello = $CustomHello + $ToastText

#Load Assemblies
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

#Build XML ToastTemplate 
[xml]$ToastTemplate = @"
<toast duration="$ToastDuration" scenario="$ToastScenario">
    <visual>
        <binding template="ToastGeneric">
            <text>$ToastTitle</text>
            <text>$CustomHello</text>
            <image placement="hero" src="$ToastImage"/>
            <group>
                <subgroup>
                    <text hint-style="body" hint-wrap="true" >$EventText</text>
                </subgroup>
            </group>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:notification.default"/>
    <actions>
        <input id="SnoozeTimer" type="selection" title="$SnoozeMessage" defaultInput="1">
            <selection id="1" content="1 Minute"/>
            <selection id="30" content="30 Minutes"/>
            <selection id="60" content="1 Hour"/>
            <selection id="120" content="2 Hours"/>
            <selection id="240" content="4 Hours"/>
        </input>
        <action activationType="system" arguments="snooze" hint-inputId="SnoozeTimer" content="$SnoozeTitle" id="test-snooze"/>
        <action arguments="dismiss" content="Dismiss" activationType="system"/>
    </actions>
</toast>
"@

#Prepare XML
$ToastXml = [Windows.Data.Xml.Dom.XmlDocument]::New()
$ToastXml.LoadXml($ToastTemplate.OuterXml)
    
#Prepare and Create Toast
$ToastMessage = [Windows.UI.Notifications.ToastNotification]::New($ToastXML)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($CustomHandlerAppID).Show($ToastMessage)
#endregion
Exit 0