<#
.SYNOPSIS
    Inventory Windows Update device settings
.DESCRIPTION
    This script is designed to be run as a Proactive Remediation. 
    Registry keys are inventoried on the device and uploaded to Log Analytics. 
    Admins are able visualise the Windows Update settings in force across their devices and understand if legacy settings or GPO's are having an undesirable effect on the device Windows Update experience
.EXAMPLE
    Invoke-WUInventory.ps1 (Required to run as System or Administrator) 
.NOTES
    FileName:    Invoke-WUInventory.ps1  
    Author:      Ben Whitmore
    Contributor: Maurice Daly
    Contact:     @byteben
    Created:     2022-10-April

    Version history:
    1.0 - (2022-04-10) Original Release
#>

#region SCRIPTVARIABLES

#Track Script Version in Log Analytics Table
$ScriptVersion = "PR003"

#Log Analytics Workspace ID
$CustomerID = ""

#Log Analytics Workspace Primary Key
$SharedKey = ""

#Custom Log Name
$LogType = "WUDevice_Settings"

# You can use an optional field to specify the timestamp from the data. If the time field is not specified, Azure Monitor assumes the time is the message ingestion time
# DO NOT DELETE THIS VARIABLE. Recommened keep this blank. 
$TimeStampField = ""

#Create Windows Update Settings Array
#Software\Policies\Microsoft\Windows\WindowsUpdate
$WUSettingsArray = @()
$WUSettingsArray += "AutoRestartDeadlinePeriodInDays"
$WUSettingsArray += "AutoRestartNotificationSchedule"
$WUSettingsArray += "AutoRestartRequiredNotificationDismissal"
$WUSettingsArray += "BranchReadinessLevel"
$WUSettingsArray += "DeferFeatureUpdates"
$WUSettingsArray += "DeferFeatureUpdatesPeriodInDays"
$WUSettingsArray += "DeferQualityUpdates"
$WUSettingsArray += "DeferQualityUpdatesPeriodInDays"
$WUSettingsArray += "DisableDualScan"
$WUSettingsArray += "DoNotConnectToWindowsUpdateInternetLocations"
$WUSettingsArray += "ElevateNonAdmins"
$WUSettingsArray += "EngagedRestartDeadline"
$WUSettingsArray += "EngagedRestartSnoozeSchedule"
$WUSettingsArray += "EngagedRestartTransitionSchedule"
$WUSettingsArray += "PauseFeatureUpdatesStartTime"
$WUSettingsArray += "PauseQualityUpdatesStartTime"
$WUSettingsArray += "ScheduleImminentRestartWarning"
$WUSettingsArray += "ScheduleRestartWarning"
$WUSettingsArray += "SetAutoRestartDeadline"
$WUSettingsArray += "SetAutoRestartNotificationConfig"
$WUSettingsArray += "SetAutoRestartNotificationDisable"
$WUSettingsArray += "SetAutoRestartRequiredNotificationDismissal"
$WUSettingsArray += "SetEDURestart"
$WUSettingsArray += "SetEngagedRestartTransitionSchedule"
$WUSettingsArray += "SetRestartWarningSchd"
$WUSettingsArray += "WUServer"
$WUSettingsArray += "WUStatusServer"

#Collect Scan Source Registry Keys
$WUSettingsArray += "SetPolicyDrivenUpdateSourceForDriverUpdates"
$WUSettingsArray += "SetPolicyDrivenUpdateSourceForFeatureUpdates"
$WUSettingsArray += "SetPolicyDrivenUpdateSourceForOtherUpdates"
$WUSettingsArray += "SetPolicyDrivenUpdateSourceForQualityUpdates"

#Software\Policies\Microsoft\Windows\WindowsUpdate\AU
$WUAUSettingsArray = @()
$WUAUSettingsArray += "AutoInstallMinorUpdates"
$WUAUSettingsArray += "EnableFeaturedSoftware"
$WUAUSettingsArray += "IncludeRecommendedUpdates"
$WUAUSettingsArray += "NoAUAsDefaultShutdownOption"
$WUAUSettingsArray += "NoAUShutdownOption"
$WUAUSettingsArray += "NoAutoRebootWithLoggedOnUsers"
$WUAUSettingsArray += "NoAutoUpdate"
$WUAUSettingsArray += "RebootRelaunchTimeout"
$WUAUSettingsArray += "RebootRelaunchTimeoutEnabled"
$WUAUSettingsArray += "RebootWarningTimeout"
$WUAUSettingsArray += "RebootWarningTimeoutEnabled"
$WUAUSettingsArray += "RescheduleWaitTime"
$WUAUSettingsArray += "RescheduleWaitTimeEnabled"

#CoManagement Capabilities for WIndows Update Workload
$CoMgmtArray = @(17, 19, 21, 23, 25, 27, 29, 31, 49, 51, 53, 55, 57, 59, 61, 63, 81, 83, 85, 87, 89, 91, 93, 95, 113, 115, 117, 119, 121, 123, 125, 127, 145, 147, 149, 151, 153, 155, 157, 159, 177, 179, 181, 183, 185, 187, 189, 191, 209, 211, 213, 215, 217, 219, 221, 223, 241, 243, 245, 247, 249, 251, 253, 255)

#endregion

#region Initialize

# Enable TLS 1.2 support 
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Get Intune DeviceID and ManagedDeviceName
if (@(Get-ChildItem HKLM:SOFTWARE\Microsoft\Enrollments\ -Recurse | Where-Object { $_.PSChildName -eq 'MS DM Server' })) {
    $MSDMServerInfo = Get-ChildItem HKLM:SOFTWARE\Microsoft\Enrollments\ -Recurse | Where-Object { $_.PSChildName -eq 'MS DM Server' }
    $ManagedDeviceInfo = Get-ItemProperty -LiteralPath "Registry::$($MSDMServerInfo)"
}
$ManagedDeviceID = $ManagedDeviceInfo.EntDMID

#Get AAD DeviceID
# Define Cloud Domain Join information registry path
$AzureADJoinInfoRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo"
		
# Retrieve the child key name that is the thumbprint of the machine certificate containing the device identifier guid
$AzureADJoinInfoThumbprint = Get-ChildItem -Path $AzureADJoinInfoRegistryKeyPath | Select-Object -ExpandProperty "PSChildName"
if (-not ($AzureADJoinInfoThumbprint -eq $null)) {
    # Retrieve the machine certificate based on thumbprint from registry key
    $AzureADJoinCertificate = Get-ChildItem -Path "Cert:\LocalMachine\My" -Recurse | Where-Object { $PSItem.Thumbprint -eq $AzureADJoinInfoThumbprint }
    if (-not ($AzureADJoinCertificate -eq $null)) {
        # Determine the device identifier from the subject name
        $AzureADDeviceID = ($AzureADJoinCertificate | Select-Object -ExpandProperty "Subject") -replace "CN=", ""
    }
}

#Get Computer Name
$DeviceName = (Get-CIMInstance -ClassName Win32_OperatingSystem -NameSpace root\cimv2).CSName

#Get OS Information
$ComputerOSVersion = (Get-CIMInstance -ClassName Win32_OperatingSystem -NameSpace root\cimv2).Version
$ComputerOSBuild = (Get-CIMInstance -ClassName Win32_OperatingSystem -NameSpace root\cimv2).BuildNumber

#Get Default AU Service
$DefaultAUService = (New-Object -ComObject "Microsoft.Update.ServiceManager").services | Where-Object { $_.IsDefaultAUService -eq 'True' } | Select-Object -ExpandProperty Name

#endregion

#region Functions

# Function to create the authorization signature
Function New-Signature ($customerID, $SharedKey, $Date, $ContentLength, $Method, $ContentType, $Resource) {
    $xHeaders = "x-ms-date:" + $Date
    $stringToHash = $Method + "`n" + $ContentLength + "`n" + $ContentType + "`n" + $xHeaders + "`n" + $Resource
	
    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)
	
    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
    return $authorization
}

# Function to create and post the request
Function Send-LogAnalyticsData($CustomerID, $SharedKey, $Body, $LogType) {
    $Method = "POST"
    $ContentType = "application/json"
    $Resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $ContentLength = $Body.Length
    $signature = New-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -ContentLength $ContentLength `
        -Method $Method `
        -ContentType $ContentType `
        -Resource $Resource
    $uri = "https://" + $CustomerID + ".ods.opinsights.azure.com" + $Resource + "?api-version=2016-04-01"
	
    #validate that payload data does not exceed limits
    if ($Body.Length -gt (31.9 * 1024 * 1024)) {
        throw ("Upload payload is too big and exceed the 32Mb limit for a single upload. Please reduce the payload size. Current payload size is: " + ($Body.Length / 1024 / 1024).ToString("#.#") + "Mb")
    }
	
    $PayLoadSize = ("Upload payload size is " + ($Body.Length / 1024).ToString("#.#") + "Kb ")
	
    $Headers = @{
        "Authorization"        = $signature;
        "Log-Type"             = $logType;
        "x-ms-date"            = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }
	
    $Response = Invoke-WebRequest -Uri $uri -Method $Method -ContentType $ContentType -Headers $Headers -Body $Body -UseBasicParsing
    $StatusMessage = "$($Response.StatusCode) : $($PayLoadSize)"
    return "$StatusMessage"
}

Function Get-RegSetting {
    param (
        [parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$RegKey,
        [string]$RegSetting
    )
    Try {
        $RegSettingValue = Get-ItemPropertyValue -Path $RegKey -Name $RegSetting -ErrorAction Stop
        If (-not ($RegSettingValue -eq $Null)) {
            Return $RegSettingValue
        }
        else {
            If ($RegSetting -in ("WUServer", "WUStatusServer")) {
                Return ""
            }
            else {
                Return 0
            }
        }
    }
    Catch {
        #Not returning caught errors
        If ($RegSetting -in ("WUServer", "WUStatusServer")) {
            Return ""
        }
        else {
            Return 0
        }
    }
}

#endregion

#region Workspace

#Build WindowsUpdate Regsitry Key Inventory
$WUSettingPayloadInventory = $Null
$WUSettingPayloadInventory = New-Object -TypeName PSObject

#Build Device Inventory
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name ScriptVersion -Value $ScriptVersion -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name DeviceName -Value $DeviceName -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name ManagedDeviceID -Value $ManagedDeviceID -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name AzureADDeviceID -Value $AzureADDeviceID -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name ComputerOSVersion -Value $ComputerOSVersion -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name ComputerOSBuild -Value $ComputerOSBuild -Force
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name DefaultAUService -Value $DefaultAUService -Force

#Get CoManagement State and Workload
$CoMgmtRegKey = "HKLM:\Software\Microsoft\CCM"
$CoMgmtSetting = "CoManagementFlags"
$CoMgmtValue = Get-RegSetting -RegKey $CoMgmtRegKey -RegSetting $CoMgmtSetting
If ($CoMgmtValue -in $CoMgmtArray) {
    $WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name CoMgmtWorkload -Value $true -Force
}
else {
    $WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name CoMgmtWorkload -Value $false -Force
}
$WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name CoMgmtValue -Value $CoMgmtValue -Force

#Get Registry Values for Software\Policies\Microsoft\Windows\WindowsUpdate
$WURegKey = "HKLM:Software\Policies\Microsoft\Windows\WindowsUpdate"
ForEach ($WUSetting in $WUSettingsArray) {
    $RegValue = Get-RegSetting -RegKey $WURegKey -RegSetting $WUSetting
    $WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name $WUSetting -Value $RegValue -Force
}

#Get Registry Values for Software\Policies\Microsoft\Windows\WindowsUpdate
$WUAURegKey = "HKLM:Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
ForEach ($WUAUSetting in $WUAUSettingsArray) {
    $RegValue = Get-RegSetting -RegKey $WUAURegKey -RegSetting $WUAUSetting
    $WUSettingPayloadInventory | Add-Member -MemberType NoteProperty -Name $WUAUSetting -Value $RegValue -Force
}

#Build Payload Array
$WUSettingPayload = $Null
$WUSettingPayload = @()
$WUSettingPayload += $WUSettingPayloadInventory

#Prepare Array for Upload
$PayloadJson = $WUSettingPayload | ConvertTo-Json

#Write upload intent to console
Write-Output "Sending Payload:"
Write-Output $PayloadJson

#Upload Data
$ResponseWUInventory = Send-LogAnalyticsData -CustomerID $CustomerID -SharedKey $SharedKey -Body ([System.Text.Encoding]::UTF8.GetBytes($PayloadJson)) -LogType $LogType
$ResponseWUInventory

#Status Report
$Date = Get-Date -Format "dd-MM HH:mm"
$OutputMessage = "InventoryDate: $Date "

if ($ResponseWUInventory) {
    if ($ResponseWUInventory -like "200*") {
        $OutputMessage = $OutputMessage + " WUInventory:OK"
    }
    else {
        $OutputMessage = $OutputMessage + " WUInventory:Fail"
    }
}
Write-Output $OutputMessage

#endregion
#>