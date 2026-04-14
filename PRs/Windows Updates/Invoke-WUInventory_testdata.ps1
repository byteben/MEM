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
    2.0 - (2026-04-01) Migrated to Logs Ingestion API (DCE/DCR) - HTTP Data Collector API deprecated Sept 2026
    1.0 - (2022-04-10) Original Release
#>

Param (
    [string]$TestDataPath = "C:\GitHub\byteben\MEM\PRs\Windows Updates\testdata.csv",
    [switch]$Interactive = $true
)

#region SCRIPTVARIABLES

$DCEEndpoint = ""
$DCRImmutableID = ""
$StreamName = "Custom-WUDevice_Settings2_CL_CL"

#endregion

#region Initialize

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#endregion

#region Functions

Function Get-AuthToken {
    if ($Interactive) {
        Try {
            Import-Module Az.Accounts -ErrorAction Stop
        }
        Catch {
            throw "Az.Accounts module not found. Install it with: Install-Module Az -Scope CurrentUser"
        }

        if (-not (Get-AzContext)) {
            Write-Output "No existing Az session found, prompting for interactive login..."
            Connect-AzAccount
        }
        else {
            Write-Output "Using existing Az session: $((Get-AzContext).Account)"
        }

        Try {
            $TokenObj = Get-AzAccessToken -ResourceUrl "https://monitor.azure.com/"
            $Token = [System.Net.NetworkCredential]::new("", $TokenObj.Token).Password
            return $Token
        }
        Catch {
            throw "Failed to obtain token via interactive login. Error: $($_.Exception.Message)"
        }
    }
    else {
        Try {
            $TokenResponse = Invoke-RestMethod `
                -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https://monitor.azure.com/" `
                -Headers @{ Metadata = "true" } `
                -Method Get `
                -ErrorAction Stop
            return $TokenResponse.access_token
        }
        Catch {
            throw "Failed to obtain Managed Identity token. Ensure the device has a Managed Identity and it has been granted Monitoring Metrics Publisher on the DCR. Error: $($_.Exception.Message)"
        }
    }
}

Function Send-LogAnalyticsData {
    param(
        [byte[]]$Body,
        [string]$Uri
    )
    
    Write-Output "Entering Send-LogAnalyticsData..."
    Write-Output "Posting to: $Uri"

    $TokenObj = Get-AzAccessToken -ResourceUrl "https://monitor.azure.com/"
    $Token = [System.Net.NetworkCredential]::new("", $TokenObj.Token).Password
    Write-Output "Token obtained successfully"

    $PayLoadSize = "Upload payload size is " + ($Body.Length / 1024).ToString("#.#") + "Kb"

    Try {
        $Response = Invoke-WebRequest `
            -Uri $Uri `
            -Method POST `
            -ContentType "application/json" `
            -Headers @{ Authorization = "Bearer $Token" } `
            -Body $Body `
            -UseBasicParsing `
            -ErrorAction Stop

        return "$($Response.StatusCode) : $PayLoadSize"
    }
    Catch {
        $ErrorBody = $_.ErrorDetails.Message
        throw "HTTP $($_.Exception.Response.StatusCode) - $ErrorBody"
    }
}

#endregion

#region Workspace

$csv = Import-Csv $TestDataPath | Select-Object `
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

$BatchSize = 100
$Batches = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $csv.Count; $i += $BatchSize) {
    $Batches.Add($csv[$i..([Math]::Min($i + $BatchSize - 1, $csv.Count - 1))])
}

Write-Output "Total rows: $($csv.Count) | Batches: $($Batches.Count)"

$BatchNumber = 0
$AllSucceeded = $true
foreach ($Batch in $Batches) {
    $BatchNumber++
    $PayloadJson = $Batch | ConvertTo-Json
    $Uri = "${DCEEndpoint}/dataCollectionRules/${DCRImmutableID}/streams/${StreamName}?api-version=2023-01-01"

    Try {
        $ResponseWUInventory = Send-LogAnalyticsData -Body ([System.Text.Encoding]::UTF8.GetBytes($PayloadJson)) -Uri $Uri
        Write-Output "Batch $BatchNumber : $ResponseWUInventory"
    }
    Catch {
        Write-Output "Batch $BatchNumber failed: $($_.Exception.Message)"
        $AllSucceeded = $false
    }
}

$Date = Get-Date -Format "dd-MM HH:mm"
if ($AllSucceeded) {
    Write-Output "InventoryDate: $Date WUInventory:OK"
}
else {
    Write-Output "InventoryDate: $Date WUInventory:Fail"
}

#endregion