<#
.SYNOPSIS
Request and generate the most common Intune reports using the Microsoft Graph API.

.Created On:   2024-09-01
.Updated On:   2025-01-21
.Created by:   Ben Whitmore @MSEndpointMgr
.Filename:     Get-IntuneReport.ps1

.DESCRIPTION
This script allows you to interactively select and fetch Intune reports using the Microsoft Graph API.
These reports are documented at https://learn.microsoft.com/en-us/mem/intune/fundamentals/reports-export-graph-available-reports
Reports are grouped by categories, enabling you to select relevant reports conveniently.
Reports are saved in the specified format (CSV by default) to a designated path.
The script supports delegated and application-based authentication flows.

.LOG FILE
The script logs its activities in CMTrace format, which is a log format used by Configuration Manager. 
The log file is saved in the specified `$SavePath` directory with the name `Get-IntuneReport.log`.

.JSON FILE STRUCTURE
The script creates JSON files in the following folder structure:
- $SavePath\<Category>_<ReportName>.json

These JSON files can be updated to reflect the filters you want to apply to the reports. Each JSON file contains the following structure:

- "reportName": the name of the report
- "category": the category of the report
- "requiredScope": the minimum scope required to call the report. If using the Connect-MgGraph cmdlet, this scope is tested to ensure the connection has at least this level of access.
- "filters": an array of available filters. Each filter contains the following structure:
  - "name": the name of the filter
  - "presence": Optional or Required. If Required, the filter must be present in the JSON with a value.
  - "operator": the operator to use when applying the filter. Can be EQ, NE, GT, GE, LT, LE, IN, or CONTAINS.
  - "value": the value to use when applying the filter.
- "properties": the properties to include in the report. Each property is a string with the name of the property. These will be sent as the Select element in the JSON body of the request.

.RESETTING JSON FILES
To reset the JSON files to their default values and remove all previous reports, pass the -OverwriteRequestBodies switch.

.REMOVING OLDER REPORTS
To remove older reports from the reports directory, pass the -CleanupOldReports switch.

.SCHEDULED TASK
To automate the execution of this script using a scheduled task, follow these steps. As this will not be an interactive session, consider using a client certificate for authentication.
For more information on how to use a certificate, visit https://msendpointmgr.com/2025/01/12/deep-diving-microsoft-graph-sdk-authentication-methods/#certificates-client-assertion

Ensure the JSON files in the $SavePath\Category_ReportName are updated with the required filters before creating the scheduled task.

1. ** Edit the script paramters

    ```powershell
    # Parameters to edit at the top of the script
    $TenantId = "your-tenant-id"
    $ClientId = "your-client-id"
    $ClientCertificateThumbprint = "your-client-certificate-thumbprint"
    $SavePath = "C:\Reports\Intune"
    $FormatChoice = "csv"
    $ReportNames = ("AllAppsList", "AppInvByDevice")
    ```

2. **Create a Scheduled Task**:
    Use the `New-ScheduledTaskAction`, `New-ScheduledTaskTrigger`, and `Register-ScheduledTask` cmdlets to create a scheduled task that runs the `Get-IntuneReport.ps1` script at a specified interval.

    ```powershell
    # Define the action to run the PowerShell script
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\Path\To\Run-IntuneReport.ps1"

    # Define the trigger to run the task daily at 2 AM
    $trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM

    # Define the principal (user) under which the task will run
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

    # Register the scheduled task
    Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "Get-IntuneReports" -Description "Scheduled task to get Intune reports daily"
    ```

3. **Verify the Scheduled Task**:
    Ensure the scheduled task is created successfully and verify its properties using the Task Scheduler GUI or the `Get-ScheduledTask` cmdlet.

    ```powershell
    # Get the details of the scheduled task
    Get-ScheduledTask -TaskName "Get-IntuneReports"
    ```

.LOCATIONS OF REPORTS:
Applications:
- AllAppsList: Found under Apps > All Apps
- AppInstallStatusAggregate: Found under Apps > Monitor > App install status
- DeviceInstallStatusByApp: Found under Apps > All Apps > Select an individual app (Required Filter: ApplicationId)
- UserInstallStatusAggregateByApp: Found under Apps > All Apps > Select an individual app (Required Filter: ApplicationId)
- DevicesByAppInv: Found under Apps > Monitor > Discovered apps > Discovered app > Export (Required Filter: ApplicationKey)
- AppInvByDevice: Found under Devices > All Devices > Device > Discovered Apps (Required Filter: DeviceId)
- AppInvAggregate: Found under Apps > Monitor > Discovered apps > Export
- AppInvRawData: Found under Apps > Monitor > Discovered apps > Export

Apps Protection:
- MAMAppProtectionStatus: Found under Apps > Monitor > App protection status > App protection report: iOS, Android
- MAMAppConfigurationStatus: Found under Apps > Monitor > App protection status > App configuration report

Cloud Attached Devices:
- ComanagedDeviceWorkloads: Found under Reports > Cloud attached devices > Reports > Co-Managed Workloads
- ComanagementEligibilityTenantAttachedDevices: Found under Reports > Cloud attached devices > Reports > Co-Management Eligibility

Device Compliance:
- DeviceCompliance: Found under Device Compliance Org
- DeviceNonCompliance: Found under Non-compliant devices

Device Management:
- Devices: Found under All devices list
- DevicesWithInventory: Found under Devices > All Devices > Export
- DeviceFailuresByFeatureUpdatePolicy: Found under Devices > Monitor > Failure for feature updates > Click on error (Required Filter: PolicyId)
- FeatureUpdatePolicyFailuresAggregate: Found under Devices > Monitor > Failure for feature updates

Endpoint Analytics:
- DeviceRunStatesByProactiveRemediation: Found under Reports > Endpoint Analytics > Proactive remediations > Select a remediation > Device status (Required Filter: PolicyId)

Endpoint Security:
- UnhealthyDefenderAgents: Found under Endpoint Security > Antivirus > Win10 Unhealthy Endpoints
- DefenderAgents: Found under Reports > Microsoft Defender > Reports > Agent Status
- ActiveMalware: Found under Endpoint Security > Antivirus > Win10 detected malware
- Malware: Found under Reports > Microsoft Defender > Reports > Detected malware
- FirewallStatus: Found under Reports > Firewall > MDM Firewall status for Windows 10 and later

Group Policy Analytics:
- GPAnalyticsSettingMigrationReadiness: Found under Reports > Group policy analytics > Reports > Group policy migration readiness

Windows Updates:
- FeatureUpdateDeviceState: Found under Reports > Windows Updates > Reports > Windows Feature Update Report (Required Filter: PolicyId)
- QualityUpdateDeviceErrorsByPolicy: Found under Devices > Monitor > Windows Expedited update failures > Select a profile (Required Filter: PolicyId)
- QualityUpdateDeviceStatusByPolicy: Found under Reports > Windows updates > Reports > Windows Expedited Update Report (Required Filter: PolicyId)

.PARAMETER SavePath
The path where the reports will be saved. Default is `$env:TEMP\IntuneReports`.

.PARAMETER FormatChoice
The format of the report. Valid formats are 'csv' and 'json'. Default is 'csv'.

.PARAMETER EndpointVersion
The Microsoft Graph API version to use. Valid endpoints are 'v1.0' and 'beta'. Default is 'beta'.

.PARAMETER ModuleNames
Module Name to connect to Graph. Default is Microsoft.Graph.Authentication.

.PARAMETER PackageProvider
If not specified, the default value NuGet is used for PackageProvider.

.PARAMETER TenantId
The Tenant ID or name to connect to.

.PARAMETER ClientId
The Application (client) ID to connect to.

.PARAMETER ClientSecret
The client secret for authentication. If a client secret is provided, the script will use the Application authentication flow.
If a client secret is not provided, the script will use the Delegated authentication flow and prompt for authorization in the browser.

.PARAMETER ClientCertificateThumbprint
The client certificate thumbprint for authentication.
See more at https://msendpointmgr.com/2025/01/12/deep-diving-microsoft-graph-sdk-authentication-methods/#certificates-client-assertion

.PARAMETER UseDeviceAuthentication
Use device authentication for Microsoft Graph API.

.PARAMETER RequiredScopes
The scopes required for Microsoft Graph API access. Default is Reports.Read.All.

.PARAMETER ModuleScope
Specifies the scope for installing the module. Default is CurrentUser.

.PARAMETER OverwriteRequestBodies
Always overwrite existing JSON file.

.PARAMETER MaxRetries
Number of retries for polling report status. Default is 10.

.PARAMETER SecondsToWait
Seconds to wait between polling report status. Default is 5.

.PARAMETER CleanupOldReports
Cleanup old reports from the reports directory.

.PARAMETER ReportNames
List of report names to run String[]. e.g. ("AllAppsList", "AppInvByDevice") or "AllAppsList" fopr a single report.

.PARAMETER LogId
Component name for logging.

.PARAMETER ResetLog
Pass this parameter to reset the log file.

.EXAMPLE
# Run the script using default settings (delegated auth, CSV format)
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id"

.EXAMPLE
# Save reports in JSON format to a specific directory
.\Get-IntuneReport.ps1 -SavePath "C:\Reports\Intune" -FormatChoice "json" -TenantId "your-tenant-id" -ClientId "your-client-id"

.EXAMPLE
# Use Application authentication flow with a client secret
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret"

.EXAMPLE
# Explicitly pass a report name and filter into the script instead of reading the report config json file
.\Get-IntuneReport.ps1 -TenantId -TenantId "your-tenant-id" -ClientId "your-client-id" -ReportNames ("AppInvByDevice", "AllAppsList")

.EXAMPLE
# Reset the log file and overwrite request bodies (return to default working directory)
Get-IntuneReport.ps1' -TenantId -TenantId "your-tenant-id" -ClientId "your-client-id" -ResetLog -OverwriteRequestBodies -CleanupOldReports
#>

param (
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 0, HelpMessage = 'The path where the reports will be saved. Default is $env:TEMP\IntuneReports')]
    [string]$SavePath = "$env:TEMP\IntuneReports",
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1, HelpMessage = 'The format of the report. Valid formats are csv and json. Default is csv')]
    [ValidateSet('csv', 'json')]
    [string]$FormatChoice = 'csv',
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 2, HelpMessage = 'The Microsoft Graph API version to use. Valid endpoints are v1.0 and beta. Default is beta')]
    [ValidateSet('v1.0', 'beta')]
    [string]$EndpointVersion = 'beta',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 3, HelpMessage = 'Module Name to connect to Graph. Default is Microsoft.Graph.Authentication')]
    [object]$ModuleNames = ('Microsoft.Graph.Authentication'),

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 4, HelpMessage = 'If not specified, the default value NuGet is used for PackageProvider')]
    [string]$PackageProvider = 'NuGet',

    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret', Position = 5, HelpMessage = 'Tenant Id or name to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientCertificateThumbprint', Position = 5, HelpMessage = 'Tenant Id or name to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'UseDeviceAuthentication', Position = 5, HelpMessage = 'Tenant Id or name to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive', Position = 5, HelpMessage = 'Tenant Id or name to connect to')]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret', Position = 6, HelpMessage = 'Client Id (App Registration) to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientCertificateThumbprint', Position = 6, HelpMessage = 'Client Id (App Registration) to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'UseDeviceAuthentication', Position = 6, HelpMessage = 'Client Id (App Registration) to connect to')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive', Position = 6, HelpMessage = 'Client Id (App Registration) to connect to')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret', Position = 7, HelpMessage = 'Client secret for authentication')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClientCertificateThumbprint', Position = 8, HelpMessage = 'Client certificate thumbprint for authentication')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientCertificateThumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = 'UseDeviceAuthentication', Position = 9, HelpMessage = 'Use device authentication for Microsoft Graph API')]
    [switch]$UseDeviceAuthentication,

    [Parameter(Mandatory = $false, Position = 10, HelpMessage = 'The scopes required for Microsoft Graph API access. Default is Reports.Read.All')]
    [string[]]$RequiredScopes = ('Reports.Read.All'),

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 11, HelpMessage = 'Specifies the scope for installing the module. Default is CurrentUser')]
    [string]$ModuleScope = 'CurrentUser',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 12, HelpMessage = 'Always overwrite existing JSON file')]
    [switch]$OverwriteRequestBodies,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 13, HelpMessage = 'Number of retries for polling report status. Default is 10')]
    [int]$MaxRetries = 10,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 14, HelpMessage = 'Seconds to wait between polling report status. Default is 5')]
    [int]$SecondsToWait = 5,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 15, HelpMessage = 'Cleanup old reports from the reports directory')]
    [switch]$CleanupOldReports,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 16, HelpMessage = 'List of report names to run')]
    [string[]]$ReportNames,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 17, HelpMessage = 'Component name for logging')]
    [string]$LogId = $($MyInvocation.MyCommand).Name,
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 18, HelpMessage = 'ResetLog: Pass this parameter to reset the log file')]
    [Switch]$ResetLog
)

#region ReportCategories
# Generate JSON for report categories
$ReportCategories = @()

# Applications
# AllAppsList: Found under Apps > All Apps
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'AllAppsList';
    Filters       = @();
    Properties    = @(
        'AppIdentifier',
        'Assigned',
        'DateCreated',
        'Description',
        'Developer',
        'ExpirationDate',
        'FeaturedApp',
        'LastModified',
        'MoreInformationURL',
        'Name',
        'Notes',
        'Owner',
        'Platform',
        'PrivacyInformationURL',
        'Publisher',
        'Status',
        'StoreURL',
        'Type',
        'Version'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# App Inv Aggregate: Found under Apps > Monitor > Discovered apps > Export
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'AppInvAggregate';
    Filters       = @();
    Properties    = @(
        'ApplicationKey',
        'ApplicationName',
        'ApplicationPublisher',
        'ApplicationShortVersion',
        'ApplicationVersion',
        'DeviceCount',
        'Platform'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# App Inv By Device: Found under Apps > Monitor > Discovered apps > Discovered app > Export
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'AppInvByDevice';
    Filters       = @(
        [PSCustomObject]@{ Name = 'DeviceId'; Presence = 'Required'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'ApplicationKey',
        'ApplicationName',
        'ApplicationPublisher',
        'ApplicationShortVersion',
        'ApplicationVersion',
        'DeviceId',
        'DeviceName',
        'OSDescription',
        'OSVersion',
        'Platform',
        'UserId',
        'EmailAddress',
        'UserName'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# App Inv Raw Data: Found under Apps > Monitor > Discovered apps > Export
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'AppInvRawData';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ApplicationName'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ApplicationPublisher'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ApplicationShortVersion'; Presence = 'Optional'; VOperator = 'eq'; alue = '' },
        [PSCustomObject]@{ Name = 'ApplicationVersion'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceId'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceName'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OSDescription'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OSVersion'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'Platform'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'UserId'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'EmailAddress'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'UserName'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'ApplicationKey',
        'ApplicationName',
        'ApplicationPublisher',
        'ApplicationShortVersion',
        'ApplicationVersion',
        'DeviceId',
        'DeviceName',
        'OSDescription',
        'OSVersion',
        'Platform',
        'UserId',
        'EmailAddress',
        'UserName'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# App Install Status Aggregate: Found under Apps > Monitor > App install status
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'AppInstallStatusAggregate';
    Filters       = @(
        [PSCustomObject]@{ Name = 'Platform'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'FailedDevicePercentage'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'ApplicationId',
        'AppVersion',
        'DisplayName',
        'FailedDeviceCount',
        'FailedDevicePercentage',
        'FailedUserCount',
        'InstalledDeviceCount',
        'InstalledUserCount',
        'NotApplicableDeviceCount',
        'NotApplicableUserCount',
        'NotInstalledDeviceCount',
        'NotInstalledUserCount',
        'PendingInstallDeviceCount',
        'PendingInstallUserCount',
        'Platform',
        'Publisher'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# Devices By App Inv: Found under Apps > Monitor > Discovered apps > Discovered app > Export
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'DevicesByAppInv';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ApplicationKey'; Presence = 'Required'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'ApplicationKey',
        'ApplicationName',
        'ApplicationPublisher',
        'ApplicationShortVersion',
        'ApplicationVersion',
        'DeviceId',
        'DeviceName',
        'OSDescription',
        'OSVersion',
        'Platform',
        'UserId',
        'EmailAddress',
        'UserName'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# Device Install Status By App: Found under Apps > All Apps > Select an individual app
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'DeviceInstallStatusByApp';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ApplicationId'; Presence = 'Required'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'AppInstallState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'HexErrorCode '; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceName',
        'UserPrincipalName',
        'UserName',
        'Platform',
        'AppVersion',
        'DeviceId',
        'AssignmentFilterIdsExist',
        'LastModifiedDateTime',
        'AppInstallState',
        'AppInstallState_loc',
        'AppInstallStateDetails',
        'AppInstallStateDetails_loc',
        'HexErrorCode'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# User Install Status Aggregate By App: Found under Apps > All Apps > Select an individual app
$ReportCategories += [PSCustomObject]@{
    Category      = 'Applications';
    ReportName    = 'UserInstallStatusAggregateByApp';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ApplicationId'; Presence = 'Required'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'UserName',
        'UserPrincipalName',
        'FailedCount',
        'InstalledCount',
        'PendingInstallCount',
        'NotInstalledCount',
        'NotApplicableCount'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# Apps Protection
# MAMAppConfigurationStatus: Found under Apps > Monitor > App protection status > App configuration report
$ReportCategories += [PSCustomObject]@{
    Category      = 'Apps Protection';
    ReportName    = 'MAMAppConfigurationStatus';
    Filters       = @();
    Properties    = @(
        'AADDeviceID',
        'AndroidPatchVersion',
        'AppAppVersion',
        'AppInstanceId',
        'DeviceHealth',
        'DeviceManufacturer',
        'DeviceModel',
        'DeviceName',
        'DeviceType',
        'MDMDeviceID',
        'Platform',
        'PlatformVersion',
        'PolicyLastSync',
        'SdkVersion',
        'UserEmail'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# MAMAppProtectionStatus: Found under Apps > Monitor > App protection status > App protection report: iOS, Android
$ReportCategories += [PSCustomObject]@{
    Category      = 'Apps Protection';
    ReportName    = 'MAMAppProtectionStatus';
    Filters       = @();
    Properties    = @(
        'AADDeviceID',
        'AndroidPatchVersion',
        'AppAppVersion',
        'AppInstanceId',
        'AppProtectionStatus',
        'ComplianceState',
        'DeviceHealth',
        'DeviceManufacturer',
        'DeviceModel',
        'DeviceName',
        'DeviceType',
        'ManagementType',
        'MDMDeviceID',
        'Platform',
        'PlatformVersion',
        'PolicyLastSync',
        'SdkVersion',
        'UserEmail'
    );
    RequiredScope = 'DeviceManagementApps.Read.All'
}

# Cloud Attached Devices
# Comanaged Device Workloads: Found under Reports > Cloud attached devices > Reports > Co-Managed Workloads
$ReportCategories += [PSCustomObject]@{
    Category      = 'Cloud Attached Devices';
    ReportName    = 'ComanagedDeviceWorkloads';
    Filters       = @(
        [PSCustomObject]@{ Name = 'CompliancePolicy'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceConfiguration'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'EndpointProtection'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ModernApps'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OfficeApps'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ResourceAccess'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'WindowsUpdateforBusiness'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceName',
        'DeviceId',
        'CompliancePolicy',
        'ResourceAccess',
        'DeviceConfiguration',
        'WindowsUpdateforBusiness',
        'EndpointProtection',
        'ModernApps',
        'OfficeApps'
    );
    RequiredScope = 'DeviceManagementManagedDevices.Read.All'
}

# Comanagement Eligibility Tenant Attached Devices: Found under Reports > Cloud attached devices > Reports > Co-Management Eligibility
$ReportCategories += [PSCustomObject]@{
    Category      = 'Cloud Attached Devices';
    ReportName    = 'ComanagementEligibilityTenantAttachedDevices';
    Filters       = @(
        [PSCustomObject]@{ Name = 'Status'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceName',
        'DeviceId',
        'Status',
        'OSDescription',
        'OSVersion'
    );
    RequiredScope = 'DeviceManagementManagedDevices.Read.All'
}

# Device Compliance
# Device Compliance: Found under Device Management > Device Compliance
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Compliance';
    ReportName    = 'DeviceCompliance';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ComplianceState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OS'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AadDeviceId',
        'ComplianceState',
        'DeviceHealthThreatLevel',
        'DeviceId',
        'DeviceName',
        'DeviceType',
        'IMEI',
        'InGracePeriodUntil',
        'IntuneDeviceId',
        'LastContact',
        'ManagementAgents',
        'OS',
        'OSDescription',
        'OSVersion',
        'OwnerType',
        'PartnerDeviceId',
        'PrimaryUser',
        'RetireAfterDatetime',
        'SerialNumber',
        'UPN',
        'UserEmail',
        'UserId',
        'UserName'
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Device Non-Compliance: Found under Device Management > Device Compliance
# 
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Compliance';
    ReportName    = 'DeviceNonCompliance';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ComplianceState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OS'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'UserId'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Device Management
# Devices: Found under Device Management > All devices list
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Management';
    ReportName    = 'Devices';
    Filters       = @(
        [PSCustomObject]@{ Name = 'CategoryName'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'CompliantState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'CreatedDate'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'EnrollmentType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'JailBroken'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'LastContact'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ManagementAgents'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ManagementState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AndroidPatchLevel',
        'CategoryId',
        'CategoryName',
        'CertExpirationDate',
        'ClientRegistrationStatus',
        'CompliantState',
        'CreatedDate',
        'DeviceId',
        'DeviceName',
        'DeviceRegistrationState',
        'DeviceState',
        'DeviceType',
        'EasAccessState',
        'EasActivationStatus',
        'EasActivationStatusString',
        'EasID',
        'EasLastSyncSuccessUtc',
        'EasStateReason',
        'EncryptionStatus',
        'EncryptionStatusString',
        'EnrolledByUser',
        'EnrollmentType',
        'EntitySource',
        'ExtendedProperties',
        'GraphDeviceIsManaged',
        'HasUnlockToken',
        'IMEI',
        'InGracePeriodUntil',
        'IsManaged',
        'JailBroken',
        'JoinType',
        'LastContact',
        'LastLoggedOnUserUPN',
        'ManagementAgents',
        'ManagementState',
        'ManagedBy',
        'ManagedDeviceName',
        'Manufacturer',
        'MDMStatus',
        'MDMWinsOverGPStartTime',
        'MEID',
        'Model',
        'OS',
        'OSDescription',
        'OSVersion',
        'OwnerType',
        'Ownership',
        'PartnerDeviceId',
        'PhoneNumber',
        'PhoneNumberE164Format',
        'PrimaryUser',
        'ReferenceId',
        'RetireAfterDatetime',
        'SCCMCoManagementFeatures',
        'SerialNumber',
        'SkuFamily',
        'StagedDeviceType',
        'StorageFree',
        'StorageTotal',
        'SubscriberCarrierNetwork',
        'SupervisedStatus',
        'SupervisedStatusString',
        'UPN',
        'UserApprovedEnrollment',
        'UserEmail',
        'UserId',
        'UserName',
        'WifiMacAddress'
    )
    RequiredScope = 'DeviceManagementManagedDevices.Read.All'
}

# Devices With Inventory: Found under Device Management > Devices > All Devices > Export
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Management';
    ReportName    = 'DevicesWithInventory';
    Filters       = @(
        [PSCustomObject]@{ Name = 'CategoryName'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'CompliantState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'CreatedDate'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'DeviceType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'EnrollmentType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'JailBroken'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'LastContact'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ManagementAgents'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ManagementState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'PartnerFeaturesBitmask'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ProcessorArchitecture'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AndroidPatchLevel',
        'CategoryName',
        'CellularTechnology',
        'CertExpirationDate',
        'CompliantState',
        'CreatedDate',
        'DeviceId',
        'DeviceName',
        'DeviceRegistrationState',
        'DeviceType',
        'EasAccessState',
        'EasActivationStatus',
        'EasID',
        'EasLastSyncSuccessUtc',
        'EasStateReason',
        'EID',
        'EnrollmentType',
        'EthernetMAC',
        'GraphDeviceIsManaged',
        'ICCID',
        'IMEI',
        'InGracePeriodUntil',
        'IsEncrypted',
        'IsManaged',
        'IsSupervised',
        'JailBroken',
        'JoinType',
        'LastContact',
        'ManagedBy',
        'ManagedDeviceName',
        'ManagementAgent',
        'ManagementAgents',
        'ManagementState',
        'Manufacturer',
        'MEID',
        'Model',
        'OSVersion',
        'OwnerType',
        'PartnerFeaturesBitmask',
        'PhoneNumber',
        'ProcessorArchitecture',
        'ReferenceId',
        'SerialNumber',
        'SkuFamily',
        'StorageFree',
        'StorageTotal',
        'SubscriberCarrierNetwork',
        'SystemManagementBIOSVersion',
        'TPMManufacturerId',
        'TPMManufacturerVersion',
        'UPN',
        'UserId',
        'UserName',
        'UserEmail',
        'WifiMacAddress',
        'WiFiIPv4Address',
        'WiFiSubnetID'
    );
    RequiredScope = 'DeviceManagementManagedDevices.Read.All'
}

# Device Failures: Found under Device Management > Device Failures
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Management';
    ReportName    = 'DeviceFailuresByFeatureUpdatePolicy';
    Filters       = @(
        [PSCustomObject]@{ Name = 'AlertMessage'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'PolicyId'; Presence = 'Required'; Value = '' },
        [PSCustomObject]@{ Name = 'RecommendedAction'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'WindowsUpdateVersion'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AADDeviceId',
        'AlertClassification',
        'AlertId',
        'AlertMessage',
        'AlertMessageData',
        'AlertMessageDescription',
        'AlertStatus',
        'AlertType',
        'Build',
        'DeviceId',
        'DeviceName',
        'EventDateTimeUTC',
        'ExtendedRecommendedAction',
        'LastUpdatedAlertStatusDateTimeUTC',
        'PolicyId',
        'PolicyName',
        'RecommendedAction',
        'ResolvedDateTimeUTC',
        'StartDateTimeUTC',
        'UPN',
        'WindowsUpdateVersion',
        'Win32ErrorCode'
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Feature Update Policy Failures: Found under Device Management > Device Failures
$ReportCategories += [PSCustomObject]@{
    Category      = 'Device Management';
    ReportName    = 'FeatureUpdatePolicyFailuresAggregate';
    Filters       = @();
    Properties    = @(
        'FeatureUpdateVersion',
        'NumberOfDevicesWithErrors',
        'PolicyId',
        'PolicyName'
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Endpoint Analytics
# Device Run States By Proactive Remediation: Found under Reports > Endpoint Analytics > Proactive remediations > Select a remediation > Device status
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Analytics';
    ReportName    = 'DeviceRunStatesByProactiveRemediation';
    Filters       = @(
        [PSCustomObject]@{ Name = 'PolicyId'; Presence = 'Required'; Value = '' },
        [PSCustomObject]@{ Name = 'DetectionStatus'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'RemediationStatus'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DetectionScriptStatus',
        'DetectionStatus',
        'DeviceId',
        'DeviceName',
        'InternalVersion',
        'ModifiedTime',
        'OSVersion',
        'PolicyId',
        'PostRemediationDetectionScriptError',
        'PostRemediationDetectionScriptOutput',
        'PreRemediationDetectionScriptError',
        'PreRemediationDetectionScriptOutput',
        'RemediationScriptErrorDetails',
        'RemediationScriptStatus',
        'RemediationStatus',
        'UniqueKustoKey',
        'UPN',
        'UserEmail',
        'UserName'
    );

    RequiredScope = 'Reports.Read.All'
}

# Endpoint Security
# Active Malware: Found under Reports > Microsoft Defender > Reports > Active Malware
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Security';
    ReportName    = 'ActiveMalware';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ExecutionState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'Severity'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'State'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceId',
        'DeviceName',
        'MalwareId',
        'MalwareName',
        'AdditionalInformationUrl',
        'Severity',
        'MalwareCategory',
        'ExecutionState',
        'State',
        'InitialDetectionDateTime',
        'LastStateChangeDateTime',
        'DetectionCount',
        'UPN',
        'UserEmail',
        'UserName'
    );

    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Defender Agents: Found under Reports > Microsoft Defender > Reports > Agent Status
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Security';
    ReportName    = 'DefenderAgents';
    Filters       = @(
        [PSCustomObject]@{ Name = 'DeviceState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'MalwareProtectionEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'NetworkInspectionSystemEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'RealTimeProtectionEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'SignatureUpdateOverdue'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    RequiredScope = 'Reports.Read.All'
}

# Firewall Status: Found under Reports > Firewall > MDM Firewall status for Windows 10 and later
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Security';
    ReportName    = 'FirewallStatus';
    Filters       = @(
        [PSCustomObject]@{ Name = 'FirewallStatus'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceName',
        'FirewallStatus',
        'FirewallStatus_loc',
        '_ManagedBy',
        'UPN'
    );
    RequiredScope = 'Reports.Read.All'
}

# Malware: Found under Reports > Microsoft Defender > Reports > Malware
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Security';
    ReportName    = 'Malware';
    Filters       = @(
        [PSCustomObject]@{ Name = 'ExecutionState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'Severity'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'State'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'DeviceId',
        'DeviceName',
        'MalwareId',
        'MalwareName',
        'AdditionalInformationUrl',
        'Severity',
        'MalwareCategory',
        'ExecutionState',
        'State',
        'InitialDetectionDateTime',
        'LastStateChangeDateTime',
        'DetectionCount',
        'UPN',
        'UserEmail',
        'UserName'
    );
    RequiredScope = 'Reports.Read.All'
}

# Unhealthy Defender Agents: Found under Endpoint Security > Antivirus > Win10 Unhealthy Endpoints
$ReportCategories += [PSCustomObject]@{
    Category      = 'Endpoint Security';
    ReportName    = 'UnhealthyDefenderAgents';
    Filters       = @(
        [PSCustomObject]@{ Name = 'DeviceState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'MalwareProtectionEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'NetworkInspectionSystemEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'RealTimeProtectionEnabled'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'SignatureUpdateOverdue'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AntiMalwareVersion',
        'CriticalFailure',
        'DeviceId',
        'DeviceName',
        'DeviceState',
        'EngineVersion',
        'FullScanOverdue',
        'FullScanRequired',
        'LastFullScanDateTime',
        'LastFullScanSignatureVersion',
        'LastQuickScanDateTime',
        'LastQuickScanSignatureVersion',
        'LastReportedDateTime',
        'MalwareProtectionEnabled',
        'NetworkInspectionSystemEnabled',
        'PendingFullScan',
        'PendingManualSteps',
        'PendingOfflineScan',
        'PendingReboot',
        'QuickScanOverdue',
        'RealTimeProtectionEnabled',
        'RebootRequired',
        'SignatureUpdateOverdue',
        'SignatureVersion',
        'UPN',
        'UserEmail',
        'UserName'
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Group Policy Analytics
# GP Analytics Setting Migration Readiness: Found under Reports > Group policy analytics > Reports > Group policy migration readiness
$ReportCategories += [PSCustomObject]@{
    Category      = 'Group Policy Analytics';
    ReportName    = 'GPAnalyticsSettingMigrationReadiness';
    Filters       = @(
        [PSCustomObject]@{ Name = 'CSPName'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'MigrationReadiness'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'ProfileType'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'CSPName',
        'MdmMapping',
        'MigrationReadiness',
        'OSVersion',
        'ProfileType',
        'Scope',
        'SettingCategory',
        'SettingName'
    );
    RequiredScope = 'DeviceManagementConfiguration.Read.All'
}

# Windows Updates
# Feature Update Device State: Found under Windows Updates > Feature updates
$ReportCategories += [PSCustomObject]@{
    Category      = 'Windows Updates';
    ReportName    = 'FeatureUpdateDeviceState';
    Filters       = @(
        [PSCustomObject]@{ Name = 'AggregateState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'LatestAlertMessage'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'PolicyId'; Presence = 'Required'; Value = '' }
    );
    Properties    = @(
        'AADDeviceId',
        'AggregateState'
        'Build',
        'CurrentDeviceUpdateStatus',
        'CurrentDeviceUpdateStatusEventDateTimeUTC',
        'CurrentDeviceUpdateSubstatus',
        'DeviceName',
        'DeviceId',
        'EventDateTimeUTC',
        'FeatureUpdateVersion',
        'LastSuccessfulDeviceUpdateStatus',
        'LastSuccessfulDeviceUpdateSubstatus',
        'LastSuccessfulDeviceUpdateStatusEventDateTimeUTC',
        'LatestAlertMessage',
        'LatestAlertMessageDescription',
        'LatestAlertRecommendedAction',
        'LatestAlertExtendedRecommendedAction',
        'LastWUScanTimeUTC',
        'OwnerType',
        'PartnerPolicyId',
        'PolicyId',
        'PolicyName',
        'UpdateCategory',
        'UPN',
        'WindowsUpdateVersion'
    );
    RequiredScope = 'Reports.Read.All'
}

# Quality Update Device Errors By Policy: Found under Windows Updates > Quality updates
$ReportCategories += [PSCustomObject]@{
    Category      = 'Windows Updates';
    ReportName    = 'QualityUpdateDeviceErrorsByPolicy';
    Filters       = @(
        [PSCustomObject]@{ Name = 'AlertMessage'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'PolicyId'; Presence = 'Required'; Value = '' }
    );
    Properties    = @(
        'AlertMessage',
        'AlertMessage_loc',
        'DeviceId',
        'DeviceName',
        'ExpediteQUReleaseDate',
        'PolicyId',
        'UPN',
        'Win32ErrorCode'
    );
    RequiredScope = 'Reports.Read.All'
}

# Quality Update Device Status By Policy: Found under Windows Updates > Quality updates
$ReportCategories += [PSCustomObject]@{
    Category      = 'Windows Updates';
    ReportName    = 'QualityUpdateDeviceStatusByPolicy';
    Filters       = @(
        [PSCustomObject]@{ Name = 'AggregateState'; Presence = 'Optional'; Operator = 'eq'; Value = '' },
        [PSCustomObject]@{ Name = 'PolicyId'; Presence = 'Required'; Value = '' },
        [PSCustomObject]@{ Name = 'OwnerType'; Presence = 'Optional'; Operator = 'eq'; Value = '' }
    );
    Properties    = @(
        'AADDeviceId',
        'AggregateState',
        'AggregateState_loc',
        'CurrentDeviceUpdateStatus',
        'CurrentDeviceUpdateStatus_loc',
        'CurrentDeviceUpdateSubstatus',
        'CurrentDeviceUpdateSubstatus_loc',
        'DeviceId',
        'DeviceName',
        'EventDateTimeUTC',
        'LastWUScanTimeUTC',
        'LatestAlertMessage',
        'LatestAlertMessage_loc',
        'OwnerType',
        'PolicyId',
        'UPN'
    );
    RequiredScope = 'Reports.Read.All'
}
#endregion

# Function to test Microsoft Graph connection
function Test-MgConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, HelpMessage = 'LogId name of the script of the calling function')]
        [string]$LogId = $($MyInvocation.MyCommand).Name,
        [Parameter(Mandatory = $false, HelpMessage = 'Scopes required for Microsoft Graph API access')]
        [string[]]$RequiredScopes,
        [Parameter(Mandatory = $false, HelpMessage = 'Test if scopes are defined')]
        [switch]$TestScopes
    )
    
    # If we don't have required scopes, set the default required scopes to create Win32 apps. This assumes the Connect-MgGraphCustom function is used outside of the New-Win32App function
    if (-not $RequiredScopes -and $TestScopes) {
        if (Test-Path variable:\global:scopes) {
            $RequiredScopes = $global:scopes
            Write-LogAndHost ("Required Scopes are defined already in global variable. Using existing required scopes: {0}" -f ($RequiredScopes -join ', ')) -ForegroundColor Green
        }
    }
    elseif (-not $RequiredScopes -and -not $TestScopes) {
        $global:scopes = @('DeviceManagementApps.ReadWrite.All')
        $RequiredScopes = $global:scopes
    }
    
    try {
        # Check if we have an active connection
        $context = Get-MgContext -ErrorAction Stop
        
        if (-not $context) {
            Write-LogAndHost "No active Microsoft Graph connection found" -ForegroundColor Yellow
            return $false
        }
    
        # Check if the required scopes are in the scopes of the active connection
        $scopes = $context.Scopes
        $missingScopes = $RequiredScopes | Where-Object { $scopes -notcontains $_ }
        
        if ($missingScopes) {
            Write-LogAndHost ("Missing required scopes: {0}" -f ($missingScopes -join ', ')) -ForegroundColor Yellow

            return $false
        }
    
        # If we get here, we have a valid connection with the required scopes
        return $true
    }
    catch {
        Write-LogAndHost ("Error while checking Microsoft Graph connection: {0}" -f $_.Exception.Message) -ForegroundColor Red

        return $false
    }
}

# Function to initialize the PowerShell module
function Initialize-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Array of module names to install')]
        [array]$Modules,
        [Parameter(Mandatory = $false, HelpMessage = 'Package provider required for module installation')]
        [string]$PackageProvider = "NuGet",
        [Parameter(Mandatory = $false, HelpMessage = 'Scope for module installation')]
        [ValidateSet('CurrentUser', 'AllUsers')]
        [string]$ModuleScope = "CurrentUser",
        [Parameter(Mandatory = $false, HelpMessage = 'Component name for logging')]
        [string]$LogId = $($MyInvocation.MyCommand).Name
    )

    begin {
        Write-LogAndHost ("Function: Initialize-Module was called for module(s): {0}" -f ($Modules -join ', ')) -LogId $LogId -ForegroundColor Cyan
    }

    process {

        try {

            # Check PackageProvider
            if (-not (Get-PackageProvider -ListAvailable -Name $PackageProvider)) {
                Write-LogAndHost ("PackageProvider not found. Installing '{0}'" -f $PackageProvider) -LogId $LogId -ForegroundColor Cyan
                Install-PackageProvider -Name $PackageProvider -ForceBootstrap -Confirm:$false
            }

            # Process each module
            foreach ($Module in $Modules) {

                if (-not (Get-Module -ListAvailable -Name $Module)) {
                    Write-LogAndHost ("Installing module '{0}' in scope '{1}'" -f $Module, $ModuleScope) -LogId $LogId -ForegroundColor Cyan
                    Install-Module -Name $Module -Scope $ModuleScope -AllowClobber -Force -Confirm:$false
                }

                if (-not (Get-Module -Name $Module)) {
                    Write-LogAndHost ("Importing module '{0}'" -f $Module) -LogId $LogId -ForegroundColor Cyan
                    
                    try {

                        # Import the module
                        Import-Module $Module
                    }
                    catch {
                        Write-LogAndHost ("Error importing module '{0}': {1}" -f $Module, $_.Exception.Message) -LogId $LogId -Severity 3

                        break
                    }
                }
                else {
                    Write-LogAndHost ("Module '{0}' already imported" -f $Module) -LogId $LogId -ForegroundColor Green
                }
            }
        }
        catch {
            Write-LogAndHost ("Module installation failed: {0}" -f $_.Exception.Message) -LogId $LogId -Severity 3

            throw
        }
    }
}

# Function to create a new report directory
function New-ReportDirectory {
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Path to save the report directory')]
        [string]$SavePath,
        [Parameter(Mandatory = $true, HelpMessage = 'Category of the report')]
        [string]$Category,
        [Parameter(Mandatory = $true, HelpMessage = 'Name of the report')]
        [string]$ReportName,
        [Parameter(Mandatory = $false, HelpMessage = 'Component name for logging')]
        [string]$LogId = $($MyInvocation.MyCommand).Name
    )
    $reportFolder = Join-Path -Path $SavePath -ChildPath ("{0}_{1}" -f $Category, $ReportName)

    if (-not (Test-Path -Path $reportFolder)) {
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
        Write-LogAndHost ("Created report directory '{0}'" -f $reportFolder) -ForegroundColor Green
    }
    return $reportFolder
}

# Function to write a log entry to a log file with CMtrace format
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, HelpMessage = 'Message to write to the log file')]
        [AllowEmptyString()]
        [String]$Message,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 1, HelpMessage = 'Location of the log file to write to')]
        [String]$LogFolder = $script:SavePath,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 2, HelpMessage = 'Name of the log file to write to. Main is the default log file')]
        [String]$Log = 'Get-IntuneReport.log',
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'LogId name of the script of the calling function')]
        [String]$LogId = $($MyInvocation.MyCommand).Name,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 3, HelpMessage = 'Severity of the log entry 1-3')]
        [ValidateSet(1, 2, 3)]
        [string]$Severity = 1,
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'The component (script name) passed as LogID to the Write-Log function including line number of invociation')]
        [string]$Component = [string]::Format('{0}:{1}', $logID, $($MyInvocation.ScriptLineNumber)),
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 4, HelpMessage = 'If specified, the log file will be reset')]
        [Switch]$ResetLogFile
    )

    Begin {
        $dateTime = Get-Date
        $date = $dateTime.ToString("MM-dd-yyyy", [Globalization.CultureInfo]::InvariantCulture)
        $time = $dateTime.ToString("HH:mm:ss.ffffff", [Globalization.CultureInfo]::InvariantCulture)
        $logToWrite = Join-Path -Path $LogFolder -ChildPath $Log
    }

    Process {
        if ($PSBoundParameters.ContainsKey('ResetLogFile')) {
            try {

                # Check if the logfile exists. We only need to reset it if it already exists
                if (Test-Path -Path $logToWrite) {

                    # Create a StreamWriter instance and open the file for writing
                    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $logToWrite
        
                    # Write an empty string to the file without the append parameter
                    $streamWriter.Write("")
        
                    # Close the StreamWriter, which also flushes the content to the file
                    $streamWriter.Close()
                    Write-LogAndHost ("Log file '{0}' wiped" -f $logToWrite) -LogId $LogId -Severity 2
                }
                else {
                    Write-LogAndHost ("Log file not found at '{0}'. Not restting log file" -f $logToWrite) -LogId $LogId -Severity 2
                }
            }
            catch {
                Write-LogAndHost -Message ("Unable to wipe log file. Error message: {0}" -f $_.Exception.Message) -LogId $LogId -Severity 3
                throw
            }
        }
            
        try {

            # Extract log object and construct format for log line entry
            foreach ($messageLine in $Message) {
                $logDetail = [string]::Format('<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">', $messageLine, $time, $date, $Component, $Context, $Severity, $PID)

                # Attempt log write
                try {
                    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $logToWrite, 'Append'
                    $streamWriter.WriteLine($logDetail)
                    $streamWriter.Close()
                }
                catch {
                    Write-Error -Message ("Unable to append log entry to '{0}' file. Error message: {1}" -f $logToWrite, $_.Exception.Message)
                    throw
                }
            }
        }
        catch [System.Exception] {
            Write-Warning -Message ("Unable to append log entry to '{0}' file" -f $logToWrite)
            throw
        }
    }
}

#Function to write a log entry to a log file and to host
function Write-LogAndHost {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, HelpMessage = 'Message to write to the log file and host')]
        [String]$Message,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 1, HelpMessage = 'Component name for logging')]
        [string]$LogId,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 2, HelpMessage = 'Foreground color for Write-LogAndHost')]
        [String]$ForegroundColor = "White",
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 4, HelpMessage = 'Create this message on a new line')]
        [Switch]$NewLine,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 3, HelpMessage = 'Severity of the log entry 1-3')]
        [ValidateSet(1, 2, 3)]
        [int]$Severity = 1
    )

    begin {

        # Set the ForegroundColor based on the severity if not already specified
        switch ($Severity) {
            2 {
                if (-not $PSBoundParameters.ContainsKey('ForegroundColor')) {
                    $ForegroundColor = "Yellow"
                }
            }
            3 {
                if (-not $PSBoundParameters.ContainsKey('ForegroundColor')) {
                    $ForegroundColor = "Red"
                }
            }
        }
    }

    process {
        
        # Call Write-Log function to write the log
        Write-Log -Message $Message -LogId $LogId -Severity $Severity

        # Write the message to the host
        if ($PSBoundParameters.ContainsKey('NewLine')) {
            Write-Host "`n$Message" -ForegroundColor $ForegroundColor
        }
        else {
            Write-Host $Message -ForegroundColor $ForegroundColor
        }
    }
}

# Set the save path globally
$script:SavePath = $SavePath

# Get the Current Date/Time in the format MM-dd-yyyy_HH-mm-ss
$dateTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

# Create base report folder if it doesn't exist
if (-not (Test-Path -Path $SavePath)) {
    New-Item -Path $SavePath -ItemType Directory -Force | Out-Null
    Write-LogAndHost "Created base directory '$SavePath'" -LogId $LogId -ForegroundColor Green 
}

# Rest the log file if the -ResetLog parameter is passed
if ($ResetLog -and (Test-Path -Path $SavePath) ) {
    Write-Log -Message $null -ResetLogFile
}


# Inform the user if existing JSON files will be overwritten
if ($OverwriteRequestBodies) {
    Write-LogAndHost ("The OverwriteRequestBodies switch was passed, existing JSON request bodies in the {0} directory will be overwritten" -f $savepath) -LogId $LogId -Severity 2
}

# Create category_reportname folders for each report and save each report category to a separate JSON file
foreach ($report in $ReportCategories) {

    # Create a new report directory
    $reportFolder = New-ReportDirectory -SavePath $SavePath -Category $report.Category -ReportName $report.ReportName

    $jsonFileName = "{0}_{1}.json" -f $report.Category, $report.ReportName
    $jsonFilePath = Join-Path -Path $reportFolder -ChildPath $jsonFileName

    if ($OverwriteRequestBodies -or -not (Test-Path -Path $jsonFilePath)) {
        $jsonContent = $report | ConvertTo-Json -Depth 10
        $jsonContent = $jsonContent -replace '(?m)^#(.*)', '/*$1*/'
        $jsonContent | Set-Content -Path $jsonFilePath -Encoding UTF8 -NoNewline
        Write-LogAndHost ("Created JSON template for report {0}" -f $report.ReportName) -LogId $LogId -ForegroundColor Green
    }
}

# Cleanup old reports if the user requested it
if ($CleanupOldReports) {

    # Regex match report names only
    $oldReports = Get-ChildItem -Path $SavePath -File  -recurse | Where-Object { $_.Name -match '^[\w-]+_\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2}\.\w+$' }

    if ($oldReports) {
        foreach ($oldReport in $oldReports) {
            Remove-Item -Path $oldReport.FullName -Recurse -Force
            Write-LogAndHost ("Removed old reports '{0}'" -f $oldReport.FullName) -LogId $LogId -ForegroundColor Green
        }
    }
    else {
        Write-LogAndHost "No old reports found to remove" -LogId $LogId -Severity 2
    }
}

# Did the user pass a report config string?
if (-not $ReportNames) {

    # Let the user select reports via Out-GridView
    Write-Host "Select the report(s) you want to fetch" -ForegroundColor Cyan
    $ReportCategoriesWithRequired = $ReportCategories | ForEach-Object {
        [PSCustomObject] [ordered]@{
            Category        = $_.Category
            ReportName      = $_.ReportName
            RequiredScope   = $_.RequiredScope
            FilterCount     = $_.Filters.Count
            RequiredFilters = ($_.Filters | Where-Object { $_.Presence -eq 'Required' }).Count
            Filters         = if ($_.Filters.Count -eq 0) { 'N/A' } else { $_.Filters }
            Properties      = $_.Properties | Sort-Object Ascending
        }
    }

    $SelectedReports = $ReportCategoriesWithRequired | 
    Sort-Object -Property Category, ReportName | 
    Out-GridView -Title "Select Intune Reports to Fetch" -PassThru
}
else {
    # Validate the provided report names against the available report names in ReportCategories
    $invalidReportNames = $ReportNames | Where-Object { $_ -notin $ReportCategories.ReportName }

    if ($invalidReportNames) {
        Write-LogAndHost ("Invalid report names provided: {0}" -f ($invalidReportNames -join ', ')) -LogId $LogId -Severity 3
        exit
    }
    else {
        $SelectedReports = $ReportCategories | Where-Object { $_.ReportName -in $ReportNames }
    }   
}

# If no reports are selected, exit
if (-not $SelectedReports) {
    Write-LogAndHost "No reports selected. Exiting..." -LogId $LogId -Severity 2
    exit
}

Write-LogAndHost ("Function: Connect-MgGraph was called with Parameter Set: {0}" -f $PSCmdlet.ParameterSetName) -LogId $LogId -ForegroundColor Cyan
Initialize-Module -Modules $ModuleNames

# Determine the authentication method based on provided parameters
if ($PSCmdlet.ParameterSetName -eq 'ClientSecret') {
    $AuthenticationMethod = 'ClientSecret'
}
elseif ($PSCmdlet.ParameterSetName -eq 'ClientCertificateThumbprint') {
    $AuthenticationMethod = 'ClientCertificateThumbprint'
}
elseif ($PSCmdlet.ParameterSetName -eq 'UseDeviceAuthentication') {
    $AuthenticationMethod = 'UseDeviceAuthentication'
}
else {
    $AuthenticationMethod = 'Interactive'
}
Write-LogAndHost ("Using authentication method: {0}" -f $AuthenticationMethod) -LogId $LogId -ForegroundColor Cyan

# First check if we already have a valid connection with required scopes
if (Test-MgConnection -RequiredScopes $RequiredScopes -TestScopes) {
    Write-LogAndHost "Using existing Microsoft Graph connection" -LogId $LogId -ForegroundColor Green
}
else {

    # If we don't have a valid connection, proceed with connection based on parameters
    $connectMgParams = [ordered]@{
        TenantId = $TenantId
    }

    switch ($AuthenticationMethod) {
        'ClientSecret' {
            $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $ClientId, $secureClientSecret
            $connectMgParams['ClientSecretCredential'] = $credential
        }
        'ClientCertificateThumbprint' {
            $connectMgParams['ClientId'] = $ClientId
            $connectMgParams['CertificateThumbprint'] = $ClientCertificateThumbprint
        }
        'UseDeviceAuthentication' {
            $connectMgParams['ClientId'] = $ClientId
            $connectMgParams['UseDeviceCode'] = $true
            $connectMgParams['Scopes'] = $RequiredScopes
        }
        'Interactive' {
            $connectMgParams['ClientId'] = $ClientId
            $connectMgParams['Scopes'] = $RequiredScopes -join ' '
        }
        default {
            Write-LogAndHost ("Unknown authentication method: {0}" -f $AuthenticationMethod) -LogId $LogId -Severity 3
            break
        }
    }

    # Convert the parameters to a string for logging
    $connectMgParamsString = 'Connect-MgGraph ' + ($connectMgParams.Keys | ForEach-Object { '-{0} {1}' -f $_, $connectMgParams.$_ }) -join ' '
    Write-LogAndHost ("Connecting to Microsoft Graph with the following parameters: {0}" -f $connectMgParamsString) -LogId $LogId -ForegroundColor Cyan
    
    try {

        # Explicitly pass the parameters to Connect-MgGraph
        if ($AuthenticationMethod -eq 'ClientSecret') {
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $connectMgParams['ClientSecretCredential'] -NoWelcome
        }
        elseif ($AuthenticationMethod -eq 'ClientCertificateThumbprint') {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $ClientCertificateThumbprint -NoWelcome
        }
        elseif ($AuthenticationMethod -eq 'UseDeviceAuthentication') {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -UseDeviceCode -Scopes $connectMgParams['Scopes'] -NoWelcome
        }
        elseif ($AuthenticationMethod -eq 'Interactive') {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Scopes $connectMgParams['Scopes'] -NoWelcome
        }
        else {
            Write-LogAndHost "Unknown authentication method: $AuthenticationMethod" -LogId $LogId -Severity 3
            break
        }
    }
    catch {
        Write-LogAndHost ("Failed to connect to Microsoft Graph: {0}" -f $_.Exception.Message) -LogId $LogId -Severity 3
    }
}
# Check if we have a valid connection with required scopes
if (Test-MgConnection -RequiredScopes $RequiredScopes) {

    Write-LogAndHost "Successfully connected to Microsoft Graph" -LogId $LogId -ForegroundColor Green
        
    # Get and display connection details
    $context = Get-MgContext
    if ($AuthenticationMethod -in @('ClientSecret', 'ClientCertificateThumbprint')) {
        Write-LogAndHost ("Connected using Client Credential Flow with application: {0}" -f $context.AppName) -LogId $LogId -ForegroundColor Green
    }
    else {
        Write-LogAndHost ("Connected using Delegated Flow as: {0}" -f $context.Account) -LogId $LogId -ForegroundColor Green
    }
    Write-LogAndHost ("Scopes defined for this Client app are: {0}" -f ($context.Scopes -join ', ')) -LogId $LogId -ForegroundColor Green
}
else {
    Write-LogAndHost "Failed to establish a valid connection with required scopes" -LogId $LogId -Severity 3
}

$ProgressPreference = 'SilentlyContinue'

# Process selected reports
$totalReports = if ($ReportConfig) { $adhocReports.Count } else { $SelectedReports.Count }
$currentReport = 0

foreach ($reportFile in $SelectedReports) {
    $currentReport++

    Write-LogAndHost ("Processing report {0} of {1}: {2}" -f $currentReport, $totalReports, $reportFile.ReportName) -LogId $LogId -ForegroundColor Cyan

    # Initialize report variable
    $report = $null
        
    # Load from JSON for GridView selections
    try {
        $reportFilePath = Join-Path -Path $SavePath -ChildPath ("{0}_{1}" -f $reportFile.Category, $reportFile.ReportName)
        $reportFile = Join-Path -Path $reportFilePath -ChildPath ("{0}_{1}.json" -f $reportFile.Category, $reportFile.ReportName)
        $report = Get-Content -Path $reportFile -Raw | ConvertFrom-Json
    }
    catch {
        Write-LogAndHost ("Failed to parse report file {0}: {1}" -f $reportFile, $_.Exception.Message) -LogId $LogId -Severity 3
        continue
    }

    $reportFolder = Join-Path -Path $SavePath -ChildPath "$($report.Category)_$($report.ReportName)"

    # Construct the request body with filters
    $body = @{
        reportName = $report.ReportName
        format     = $FormatChoice
        filter     = ""
        select     = $report.Properties
    }

    # Prompt for required filters
    $requiredFilters = @($report.Filters) | ForEach-Object { $_ | Where-Object { $_.Presence -eq 'Required' } }

    foreach ($filter in $requiredFilters) {
        $dontContinue = $false

        try {

            # Prompt for filter value if not provided in the JSON
            if ([string]::IsNullOrEmpty($filter.Value)) {
                $filterValue = Read-Host -Prompt "Enter value for required filter '$($filter.Name)'"
            }
            else {
                $filterValue = $filter.Value
            }

            # Validate the required filter value
            if ([string]::IsNullOrWhiteSpace($filterValue)) {
                Write-LogAndHost "No value supplied for required filter '$($filter.Name)'. Skipping report request." -LogId $LogId -Severity 2
                $dontContinue = $true
                break
            }
            else {
                Write-LogAndHost ("Setting filter value for '{0}' to '{1}'" -f $filter.Name, $filterValue) -LogId $LogId -ForegroundColor Cyan
            
                # Build filter string
                if ([string]::IsNullOrEmpty($body.filter)) {
                    $body.filter = "({0} {1} '{2}')" -f $filter.Name, $filter.Operator, $filterValue
                }
                else {
                    if ($body.filter.StartsWith("(")) {
                        $body.filter = $body.filter.Substring(0, $body.filter.Length - 1)
                    }

                    $body.filter += " and {0} {1} '{2}'" -f $filter.Name, $filter.Operator, $filterValue

                    if (-not $body.filter.EndsWith(")")) {
                        $body.filter += ")"
                    }
                }
            }
        }
        catch {
            Write-LogAndHost ("Error setting filter value: {0}" -f $_) -LogId $LogId -Severity 3
            $dontContinue = $true
            return
        }
    }

    if ($dontContinue) {
        Write-LogAndHost "Required filter values missing. Skipping report request." -LogId $LogId -Severity 2
        return
    }
    else {

        # Convert the body to JSON
        $bodyJson = $body | ConvertTo-Json -Depth 10
        Write-LogAndHost ("Body to post: {0}" -f $bodyJson) -LogId $LogId -ForegroundColor Cyan

        try {

            $uri = "https://graph.microsoft.com/$EndpointVersion/deviceManagement/reports/exportJobs"
            Write-LogAndHost ("Posting request to {0}" -f $uri) -LogId $LogId -ForegroundColor Cyan
            $response = Invoke-MgGraphRequest -Uri $uri -Method Post -Body $bodyJson -ContentType "application/json"

            $jobId = $response.id
            $pollingUri = "https://graph.microsoft.com/$EndpointVersion/deviceManagement/reports/exportJobs/$jobId"
            Write-LogAndHost ("Report request {0} submitted successfully. Polling URI: {1}" -f $report.ReportName, $pollingUri) -LogId $LogId -ForegroundColor Cyan

            $progressPreference = 'Continue'

            # Set the start time and the maximum number of attempts
            $startTime = Get-Date
            $attempts = 1
            $lastStatusCheckTime = $startTime

            while ($attempts -lt $maxRetries) {
                $currentTime = Get-Date
                $timeElapsed = ($currentTime - $startTime).TotalSeconds
                $timeElapsedString = [TimeSpan]::FromSeconds($timeElapsed).ToString("hh\:mm\:ss")
    
                # Update host output every 1 second
                Write-Host ("`rElapsedTime: {0}" -f $timeElapsedString) -NoNewline -ForegroundColor Yellow
    
                $progressPercent = [math]::Round(($attempts / $maxRetries) * 100, 0)
                Write-Progress -Activity "Polling Report Status" -Status "Attempt $attempts of $maxRetries" -PercentComplete $progressPercent
    
                # Only check status every 6 seconds
                if (($currentTime - $lastStatusCheckTime).TotalSeconds -ge 6) {
                    $statusResponse = Invoke-MgGraphRequest -Uri $pollingUri -Method Get
                    $lastStatusCheckTime = $currentTime

                    if ($statusResponse.status -eq "completed") {
                        Write-LogAndHost ("`nReport {0} is ready for download" -f $report.ReportName) -LogId $LogId -ForegroundColor Cyan
                        $downloadUri = $statusResponse.url
                        Write-LogAndHost ("Download URI: {0}" -f $downloadUri) -LogId $LogId -ForegroundColor Green
                        $zipFilePath = "$reportFolder\$($report.ReportName).zip"
                        Write-LogAndHost ("Downloading report {0} to {1}" -f $report.ReportName, $zipFilePath) -LogId $LogId -ForegroundColor Cyan
                        Invoke-MgGraphRequest -Uri $downloadUri -OutputFilePath $zipFilePath -Method Get 
                        Write-LogAndHost ("Report {0} downloaded successfully to {1}" -f $report.ReportName, $zipFilePath) -LogId $LogId -ForegroundColor Green
                        Write-LogAndHost ("Report {0} took {1} seconds to generate and download" -f $report.ReportName, $timeElapsed) -LogId $LogId -ForegroundColor Cyan

                        # Extract the zip file to the same directory
                        try {
                            $extractPath = $reportFolder
                            Write-LogAndHost ("Extracting report {0} to {1}" -f $report.ReportName, $extractPath) -LogId $LogId -ForegroundColor Cyan
                            $extractedFiles = Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force -PassThru
                        }
                        catch {
                            Write-LogAndHost ("Error extracting report {0} to {1}: {2}" -f $report.ReportName, $extractPath, $_.Exception.Message) -LogId $LogId -Severity 3
                            break
                        }
                        Write-LogAndHost ("Report {0} extracted successfully to {1}" -f $report.ReportName, $extractPath) -LogId $LogId -ForegroundColor Green
                    
                        # Rename the extracted files
                        foreach ($file in $extractedFiles) {
                            $fileExtension = [System.IO.Path]::GetExtension($file.FullName)

                            # Clean up the report name to remove any illegal characters
                            $safeReportName = $report.ReportName -replace '[\\/:*?"<>|]', '_'
                            $newFileName = "{0}_{1}{2}" -f $safeReportName, $dateTime, $fileExtension
    
                            # Use resolved full paths
                            $sourceFullPath = [System.IO.Path]::GetFullPath($file.FullName)
                            $targetFullPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($extractPath, $newFileName))
    
                            Write-LogAndHost ("Renaming file {0} to {1}" -f $sourceFullPath, $targetFullPath) -LogId $LogId -ForegroundColor Cyan

                            try {

                                # Ensure target directory exists
                                $targetDir = [System.IO.Path]::GetDirectoryName($targetFullPath)
                                if (-not (Test-Path -Path $targetDir)) {
                                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                                }
        
                                # Move-Item instead of Rename-Item for better cross-device support
                                Move-Item -LiteralPath $sourceFullPath -Destination $targetFullPath -Force
                                Write-LogAndHost ("Successfully renamed file to {0}" -f $newFileName) -LogId $LogId -ForegroundColor Green
                            }
                            catch {
                                Write-LogAndHost ("Failed to rename file {0} to {1}: {2}" -f $sourceFullPath, $targetFullPath, $_.Exception.Message) -LogId $LogId -Severity 3
                                continue
                            }
                        }

                        # Delete the zip file
                        Write-LogAndHost ("Deleting zip file {0}" -f $zipFilePath) -LogId $LogId -Severity 2
                        Remove-Item -Path $zipFilePath
                        break
                    }
                    $attempts++
                }

                # Small sleep to prevent high CPU usage
                Start-Sleep -Milliseconds 500

                if ($attempts -ge $maxRetries) {
                    Write-LogAndHost ("Failed to fetch report {0} after {1} attempts. Please try again later." -f $report.ReportName, $maxRetries) -LogId $LogId -Severity 3
                }
            }
        }
        catch {
            Write-LogAndHost ("Failed to fetch report {0}: {1}" -f $report.ReportName, $_.Exception) -LogId $LogId -Severity 3
            return
        }
    }
}

# Disconnect from Microsoft Graph
if (Get-MgContext) {
    $userInput = Read-Host -Prompt "Do you want to disconnect from Microsoft Graph? (y/n, default: n)"

    if ($userInput -eq '') {
        $userInput = 'n'
    }
    switch ($userInput.ToLower()) {
        "n" {
            Write-LogAndHost -Message "Leaving Microsoft Graph session open" -LogId $LogId -ForegroundColor Cyan
            break
        }
        "y" {
            Write-LogAndHost "Disconnecting from Microsoft Graph" -LogId $LogId -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
        }
        default {
            Write-LogAndHost "Invalid input. Please type 'y' or 'n'." -LogId $LogId -Severity 2
        }
    }
}
else {
    Write-LogAndHost "No active Microsoft Graph connection found" -LogId $LogId -Severity 2
}