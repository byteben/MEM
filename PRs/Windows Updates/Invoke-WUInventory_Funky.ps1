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
$ScriptVersion = "PR004"

#FunctionAppURI for Log Upload
$AzureFunctionURL = "https://fn-mycontosoinventoryfunction.azurewebsites.net/api/LogCollectorAPI"

#Custom Log Name
$CustomLogName = "WUDevice_Settings"

#Date
$Date = (Get-Date)

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
function Get-AzureADDeviceID {
    <#
    .SYNOPSIS
        Get the Azure AD device ID from the local device.
    
    .DESCRIPTION
        Get the Azure AD device ID from the local device.
    
    .NOTES
        Author:      Nickolaj Andersen
        Contact:     @NickolajA
        Created:     2021-05-26
        Updated:     2021-05-26
    
        Version history:
        1.0.0 - (2021-05-26) Function created
    #>
    Process {
        # Define Cloud Domain Join information registry path
        $AzureADJoinInfoRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo"
		
        # Retrieve the child key name that is the thumbprint of the machine certificate containing the device identifier guid
        $AzureADJoinInfoThumbprint = Get-ChildItem -Path $AzureADJoinInfoRegistryKeyPath | Select-Object -ExpandProperty "PSChildName"
        if ($AzureADJoinInfoThumbprint -ne $null) {
            # Retrieve the machine certificate based on thumbprint from registry key
            $AzureADJoinCertificate = Get-ChildItem -Path "Cert:\LocalMachine\My" -Recurse | Where-Object { $PSItem.Thumbprint -eq $AzureADJoinInfoThumbprint }
            if ($AzureADJoinCertificate -ne $null) {
                # Determine the device identifier from the subject name
                $AzureADDeviceID = ($AzureADJoinCertificate | Select-Object -ExpandProperty "Subject") -replace "CN=", ""
                # Handle return value
                return $AzureADDeviceID
            }
        }
    }
} #endfunction 
function Get-AzureADJoinDate {
    <#
    .SYNOPSIS
        Get the Azure AD device ID from the local device.
    
    .DESCRIPTION
        Get the Azure AD device ID from the local device.
    
    .NOTES
        Author:      Nickolaj Andersen
        Contact:     @NickolajA
        Created:     2021-05-26
        Updated:     2021-05-26
    
        Version history:
        1.0.0 - (2021-05-26) Function created
    #>
    Process {
        # Define Cloud Domain Join information registry path
        $AzureADJoinInfoRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo"
		
        # Retrieve the child key name that is the thumbprint of the machine certificate containing the device identifier guid
        $AzureADJoinInfoThumbprint = Get-ChildItem -Path $AzureADJoinInfoRegistryKeyPath | Select-Object -ExpandProperty "PSChildName"
        if ($AzureADJoinInfoThumbprint -ne $null) {
            # Retrieve the machine certificate based on thumbprint from registry key
            $AzureADJoinCertificate = Get-ChildItem -Path "Cert:\LocalMachine\My" -Recurse | Where-Object { $PSItem.Thumbprint -eq $AzureADJoinInfoThumbprint }
            if ($AzureADJoinCertificate -ne $null) {
                # Determine the device identifier from the subject name
                $AzureADJoinDate = ($AzureADJoinCertificate | Select-Object -ExpandProperty "NotBefore") 
                # Handle return value
                return $AzureADJoinDate
            }
        }
    }
} #endfunction 
#Function to get AzureAD TenantID
function Get-AzureADTenantID {
    # Cloud Join information registry path
    $AzureADTenantInfoRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo"
    # Retrieve the child key name that is the tenant id for AzureAD
    $AzureADTenantID = Get-ChildItem -Path $AzureADTenantInfoRegistryKeyPath | Select-Object -ExpandProperty "PSChildName"
    return $AzureADTenantID
}                          
#endregion functions


#region Workspace
#Get Common data for validation in Azure Function: 
$AzureADDeviceID = Get-AzureADDeviceID
$AzureADTenantID = Get-AzureADTenantID

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

#Randomize over 50 minutes to spread load on Azure Function - disabled on date of enrollment (Disabled in sample - Enable only in larger environments)
$JoinDate = Get-AzureADJoinDate
$DelayDate = $JoinDate.AddDays(1)
$CompareDate = ($DelayDate - $JoinDate)
if ($CompareDate.Days -ge 1) {
    #Write-Output "Randomzing execution time"
    #$ExecuteInSeconds = (Get-Random -Maximum 3000 -Minimum 1)
    #Start-Sleep -Seconds $ExecuteInSeconds
}
#Start sending logs
$date = Get-Date -Format "dd-MM HH:mm"
$OutputMessage = "InventoryDate:$date "

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")

# Add arrays of logs into payload array 
$LogPayLoad = New-Object -TypeName PSObject 
$LogPayLoad | Add-Member -NotePropertyMembers @{$CustomLogName = $WUSettingPayload }

# Construct main payload to send to LogCollectorAPI // IMPORTANT // KEEP AND DO NOT CHANGE THIS
$MainPayLoad = [PSCustomObject]@{
    AzureADTenantID = $AzureADTenantID
    AzureADDeviceID = $AzureADDeviceID
    LogPayloads     = $LogPayLoad
}
$MainPayLoadJson = $MainPayLoad | ConvertTo-Json -Depth 9	

#Write upload intent to console
Write-Output "Sending Payload:"
Write-Output $MainPayLoadJson

# Sending data to API
try {
    $ResponseInventory = Invoke-RestMethod $AzureFunctionURL -Method 'POST' -Headers $headers -Body $MainPayLoadJson
    $OutputMessage = $OutPutMessage + "Inventory:OK " + $ResponseInventory
} 
catch {
    $ResponseInventory = "Error Code: $($_.Exception.Response.StatusCode.value__)"
    $ResponseMessage = $_.Exception.Message
    $OutputMessage = $OutPutMessage + "Inventory:Fail " + $ResponseInventory + $ResponseMessage
}

# Check status and report to Proactive Remediations
if ($ResponseInventory -match "200") {
    Write-Output $OutputMessage
    Exit 0
}
else {
    Write-Output "Error: $($ResponseInventory), Message: $($ResponseMessage)"
    Exit 1
}

#endregion
#>