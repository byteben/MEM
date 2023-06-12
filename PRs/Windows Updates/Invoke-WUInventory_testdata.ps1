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
Param (
    [Parameter(Mandatory = $true)]
    [string]$testdatapath 
)

#region SCRIPTVARIABLES

#Log Analytics Workspace ID
$CustomerID = "13ef5887-f37c-432d-be47-17c3aa40fc55"

#Log Analytics Workspace Primary Key
$SharedKey = "+f9wKsv3BLWX1nlU/ywFQIShFeF8CUyPTECHE+Wc0LW3DwcONLU5UYpcXqIFn3/cGK2obsCPw9jvCebJJlEZCg=="

#Custom Log Name
$LogType = "WUDevice_Settings"

# You can use an optional field to specify the timestamp from the data. If the time field is not specified, Azure Monitor assumes the time is the message ingestion time
# DO NOT DELETE THIS VARIABLE. Recommened keep this blank. 
$TimeStampField = ""

#region Initialize

# Enable TLS 1.2 support 
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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

#region Workspace

#Import test data
$csv = Import-Csv $testdatapath | Select-Object `
    ScriptVersion, `
    DeviceName, `
    ManagedDeviceID, `
    AzureADDeviceID, `
    ComputerOSVersion, `
    ComputerOSBuild, `
    DefaultAUService, `
@{Name = "CoMgmtWorkload"; Expression = { [bool]$_.CoMgmtWorkload } }, `
@{Name = "CoMgmtValue"; Expression = { [int]$_.CoMgmtValue } }, `
@{Name = "AutoInstallMinorUpdates"; Expression = { [int]$_.AutoInstallMinorUpdates } }, `
@{Name = "AutoRestartDeadlinePeriodInDays"; Expression = { [int]$_.AutoRestartDeadlinePeriodInDays } }, `
@{Name = "AutoRestartNotificationSchedule"; Expression = { [int]$_.AutoRestartNotificationSchedule } }, `
@{Name = "AutoRestartRequiredNotificationDismissal"; Expression = { [int]$_.AutoRestartRequiredNotificationDismissal } }, `
@{Name = "BranchReadinessLevel"; Expression = { [int]$_.BranchReadinessLevel } }, `
@{Name = "DeferFeatureUpdates"; Expression = { [int]$_.DeferFeatureUpdates } }, `
@{Name = "DeferFeatureUpdatesPeriodInDays"; Expression = { [int]$_.DeferFeatureUpdatesPeriodInDays } }, `
@{Name = "DeferQualityUpdates"; Expression = { [int]$_.DeferQualityUpdates } }, `
@{Name = "DeferQualityUpdatesPeriodInDays"; Expression = { [int]$_.DeferQualityUpdatesPeriodInDays } }, `
@{Name = "DisableDualScan"; Expression = { [int]$_.DisableDualScan } }, `
@{Name = "DoNotConnectToWindowsUpdateInternetLocations"; Expression = { [int]$_.DoNotConnectToWindowsUpdateInternetLocations } }, `
@{Name = "ElevateNonAdmins"; Expression = { [int]$_.ElevateNonAdmins } }, `
@{Name = "EnableFeaturedSoftware"; Expression = { [int]$_.EnableFeaturedSoftware } }, `
@{Name = "EngagedRestartDeadline"; Expression = { [int]$_.EngagedRestartDeadline } }, `
@{Name = "EngagedRestartSnoozeSchedule"; Expression = { [int]$_.EngagedRestartSnoozeSchedule } }, `
@{Name = "EngagedRestartTransitionSchedule"; Expression = { [int]$_.EngagedRestartTransitionSchedule } }, `
@{Name = "IncludeRecommendedUpdates"; Expression = { [int]$_.IncludeRecommendedUpdates } }, `
@{Name = "NoAUAsDefaultShutdownOption"; Expression = { [int]$_.NoAUAsDefaultShutdownOption } }, `
@{Name = "NoAUShutdownOption"; Expression = { [int]$_.NoAUShutdownOption } }, `
@{Name = "NoAutoRebootWithLoggedOnUsers"; Expression = { [int]$_.NoAutoRebootWithLoggedOnUsers } }, `
@{Name = "NoAutoUpdate"; Expression = { [int]$_.NoAutoUpdate } }, `
@{Name = "PauseFeatureUpdatesStartTime"; Expression = { [int]$_.PauseFeatureUpdatesStartTime } }, `
@{Name = "PauseQualityUpdatesStartTime"; Expression = { [int]$_.PauseQualityUpdatesStartTime } }, `
@{Name = "RebootRelaunchTimeout"; Expression = { [int]$_.RebootRelaunchTimeout } }, `
@{Name = "RebootRelaunchTimeoutEnabled"; Expression = { [int]$_.RebootRelaunchTimeoutEnabled } }, `
@{Name = "RebootWarningTimeout"; Expression = { [int]$_.RebootWarningTimeout } }, `
@{Name = "RebootWarningTimeoutEnabled"; Expression = { [int]$_.RebootWarningTimeoutEnabled } }, `
@{Name = "RescheduleWaitTime"; Expression = { [int]$_.RescheduleWaitTime } }, `
@{Name = "RescheduleWaitTimeEnabled"; Expression = { [int]$_.RescheduleWaitTimeEnabled } }, `
@{Name = "ScheduleImminentRestartWarning"; Expression = { [int]$_.ScheduleImminentRestartWarning } }, `
@{Name = "ScheduleRestartWarning"; Expression = { [int]$_.ScheduleRestartWarning } }, `
@{Name = "SetAutoRestartDeadline"; Expression = { [int]$_.SetAutoRestartDeadline } }, `
@{Name = "SetAutoRestartNotificationConfig"; Expression = { [int]$_.SetAutoRestartNotificationConfig } }, `
@{Name = "SetAutoRestartNotificationDisable"; Expression = { [int]$_.SetAutoRestartNotificationDisable } }, `
@{Name = "SetAutoRestartRequiredNotificationDismissal"; Expression = { [int]$_.SetAutoRestartRequiredNotificationDismissal } }, `
@{Name = "SetEDURestart"; Expression = { [int]$_.SetEDURestart } }, `
@{Name = "SetEngagedRestartTransitionSchedule"; Expression = { [int]$_.SetEngagedRestartTransitionSchedule } }, `
@{Name = "SetPolicyDrivenUpdateSourceForDriverUpdates"; Expression = { [int]$_.SetPolicyDrivenUpdateSourceForDriverUpdates } }, `
@{Name = "SetPolicyDrivenUpdateSourceForFeatureUpdates"; Expression = { [int]$_.SetPolicyDrivenUpdateSourceForFeatureUpdates } }, `
@{Name = "SetPolicyDrivenUpdateSourceForOtherUpdates"; Expression = { [int]$_.SetPolicyDrivenUpdateSourceForOtherUpdates } }, `
@{Name = "SetPolicyDrivenUpdateSourceForQualityUpdates"; Expression = { [int]$_.SetPolicyDrivenUpdateSourceForQualityUpdates } }, `
@{Name = "SetRestartWarningSchd"; Expression = { [int]$_.SetRestartWarningSchd } }, `
    WUServer, `
    WUStatusServer


#Prepare Array for Upload
$PayloadJson = $csv | ConvertTo-Json

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