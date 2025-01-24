# Get-IntuneReport

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue?logo=powershell)  
![Status](https://img.shields.io/badge/Status-Preview-yellow)

## 📝 Synopsis

The Get-IntuneReport script allows you to generate and download reports from Microsoft Intune. It supports various report types with filtering capabilities and authentication methods.  

![Out-GridView](https://byteben.com/bb/Downloads/GitHub/Get-IntuneReport_SelectReports.png)

## 📋 Requirements

- PowerShell 5.1 or later  
- Microsoft.Graph.Authentication module  
- Entra ID App Registration with appropriate permissions  
- Internet connectivity to access Microsoft Graph API  

## 📖 Description

This script allows you to interactively select and fetch Intune reports using the Microsoft Graph API. These reports are documented at [Microsoft Learn](https://learn.microsoft.com/en-us/mem/intune/fundamentals/reports-export-graph-available-reports).  
Reports are grouped by categories, enabling you to select relevant reports conveniently. Reports are saved in the specified format (CSV by default) to the designated path ($SavePath).  
The script supports delegated and application-based authentication flows.  

## 📄 Log File

The script logs its activities in CMTrace format, which is a log format used by Configuration Manager.  
The log file is saved in the specified `$SavePath` directory with the name `Get-IntuneReport.log`.  

![Log File](https://byteben.com/bb/Downloads/GitHub/Get-IntuneReport_Log.png)

## 📂 JSON File Structure

The script creates JSON files in the following folder structure:

- `$SavePath\<Category>_<ReportName>.json`

These JSON files can be updated to reflect the filters you want to apply to the reports. Each JSON file contains the following structure:

- `reportName`: the name of the report
- `category`: the category of the report
- `requiredScope`: the minimum scope required to call the report. If using the `Connect-MgGraph` cmdlet, this scope is tested to ensure the connection has at least this level of access.
- `filters`: an array of available filters. Each filter contains the following structure:
  - `name`: the name of the filter
  - `presence`: Optional or Required. If Required, the filter must be present in the JSON with a value.
  - `operator`: the operator to use when applying the filter. Can be EQ, NE, GT, GE, LT, LE, IN, or CONTAINS.
  - `value`: the value to use when applying the filter.
- `properties`: the properties to include in the report. Each property is a string with the name of the property. These will be sent as the Select element in the JSON body of the request. You can remove properties that you do not want returned in the report.

![JSON Structure](https://byteben.com/bb/Downloads/GitHub/Get-IntuneReport_JSONExample.png)  

The image below indicates the user is prompted for a filter value if it is required and not supplied in the JSON.  

![Required Filters](https://byteben.com/bb/Downloads/GitHub/Get-IntuneReport_RequiredFilter.png)

## 🔄 Resetting JSON Files

To reset the JSON files to their default values, pass the `-OverwriteRequestBodies` switch.

## 🗑️ Removing Older Reports

To remove older reports from the reports directory, pass the `-CleanupOldReports` switch.

## 🚀 Script Usage

To use the `Get-IntuneReport` script, follow the examples below. These examples demonstrate how to run the script with various parameters and options.  
By default, all properties are returned in the report. To customize the properties returned and the filters used in the request, edit the JSON files in the `$SavePath\Category_ReportName` directory.  

### Quick Start Run 
`Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id"` to connect interactively to the Microsoft Graph API and display an interactive menu to select reports you wish to request.  

The example below indicates the host output when running the script and selecting reports 2 reports.  

![Example Host Output](https://byteben.com/bb/Downloads/GitHub/Get-IntuneReport_ExampleRun.png)



## ⏲️ Scheduled Task

To automate the execution of this script using a scheduled task, follow these steps. As this will not be an interactive session, consider using a client certificate for authentication. For more information on how to use a certificate, visit [Deep Diving Microsoft Graph SDK Authentication Methods](https://msendpointmgr.com/2025/01/12/deep-diving-microsoft-graph-sdk-authentication-methods/#certificates-client-assertion).

Ensure the JSON files in the `$SavePath\Category_ReportName` are updated with the required filters before creating the scheduled task.

1. **Edit the script parameters**:  
        Update the script parameters in the `Run-IntuneReport.ps1` script to match your environment. The following parameters are required:  

```yaml
$TenantId = "your-tenant-id"
$ClientId = "your-client-id"
$ClientCertificateThumbprint = "your-client-certificate-thumbprint"
$SavePath = "C:\Reports\Intune"
$FormatChoice = "csv"
$ReportNames = ("AllAppsList", "AppInvByDevice")
 ```

2. **Create a Scheduled Task**:  
        Use the `New-ScheduledTaskAction`, `New-ScheduledTaskTrigger`, and `Register-ScheduledTask` cmdlets to create a scheduled task that runs the `Run-IntuneReport.ps1` script at a specified interval.  

```yaml
# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\Path\To\Get-IntuneReport.ps1"

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

## 📍 Locations of Reports in the Intune Admin Centre
For more information on available reports, visit [Microsoft Learn](https://learn.microsoft.com/en-us/mem/intune/fundamentals/reports-export-graph-available-reports).  

### Applications

- **AllAppsList**: Found under Apps > All Apps
- **AppInstallStatusAggregate**: Found under Apps > Monitor > App install status
- **DeviceInstallStatusByApp**: Found under Apps > All Apps > Select an individual app (Required Filter: ApplicationId)
- **UserInstallStatusAggregateByApp**: Found under Apps > All Apps > Select an individual app (Required Filter: ApplicationId)
- **DevicesByAppInv**: Found under Apps > Monitor > Discovered apps > Discovered app > Export (Required Filter: ApplicationKey)
- **AppInvByDevice**: Found under Devices > All Devices > Device > Discovered Apps (Required Filter: DeviceId)
- **AppInvAggregate**: Found under Apps > Monitor > Discovered apps > Export
- **AppInvRawData**: Found under Apps > Monitor > Discovered apps > Export

### Apps Protection

- **MAMAppProtectionStatus**: Found under Apps > Monitor > App protection status > App protection report: iOS, Android
- **MAMAppConfigurationStatus**: Found under Apps > Monitor > App protection status > App configuration report

### Cloud Attached Devices

- **ComanagedDeviceWorkloads**: Found under Reports > Cloud attached devices > Reports > Co-Managed Workloads
- **ComanagementEligibilityTenantAttachedDevices**: Found under Reports > Cloud attached devices > Reports > Co-Management Eligibility

### Device Compliance

- **DeviceCompliance**: Found under Device Compliance Org
- **DeviceNonCompliance**: Found under Non-compliant devices

### Device Management

- **Devices**: Found under All devices list
- **DevicesWithInventory**: Found under Devices > All Devices > Export
- **DeviceFailuresByFeatureUpdatePolicy**: Found under Devices > Monitor > Failure for feature updates > Click on error (Required Filter: PolicyId)
- **FeatureUpdatePolicyFailuresAggregate**: Found under Devices > Monitor > Failure for feature updates

### Endpoint Analytics

- **DeviceRunStatesByProactiveRemediation**: Found under Reports > Endpoint Analytics > Proactive remediations > Select a remediation > Device status (Required Filter: PolicyId)

### Endpoint Security

- **UnhealthyDefenderAgents**: Found under Endpoint Security > Antivirus > Win10 Unhealthy Endpoints
- **DefenderAgents**: Found under Reports > Microsoft Defender > Reports > Agent Status
- **ActiveMalware**: Found under Endpoint Security > Antivirus > Win10 detected malware
- **Malware**: Found under Reports > Microsoft Defender > Reports > Detected malware
- **FirewallStatus**: Found under Reports > Firewall > MDM Firewall status for Windows 10 and later

### Group Policy Analytics

- **GPAnalyticsSettingMigrationReadiness**: Found under Reports > Group policy analytics > Reports > Group policy migration readiness

### Windows Updates

- **FeatureUpdateDeviceState**: Found under Reports > Windows Updates > Reports > Windows Feature Update Report (Required Filter: PolicyId)
- **QualityUpdateDeviceErrorsByPolicy**: Found under Devices > Monitor > Windows Expedited update failures > Select a profile (Required Filter: PolicyId)
- **QualityUpdateDeviceStatusByPolicy**: Found under Reports > Windows updates > Reports > Windows Expedited Update Report (Required Filter: PolicyId)

### Example 1: Basic Usage

This example shows how to run the script with the default parameters. The reports will be saved in the default path (`$env:TEMP\IntuneReports`) in CSV format.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret"
```

### Example 2: Specify Save Path and Format

This example demonstrates how to specify a custom save path and output format (JSON).

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret" -SavePath "C:\CustomReports\Intune" -FormatChoice "json"
```

### Example 3: Use Client Certificate for Authentication

This example shows how to use a client certificate for authentication.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientCertificateThumbprint "your-client-certificate-thumbprint"
```

### Example 4: Run Specific Reports

This example demonstrates how to run specific reports by specifying their names.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret" -ReportNames "AllAppsList", "AppInvByDevice"
```

### Example 5: Overwrite Existing JSON Files

This example shows how to overwrite existing JSON files with default values. This is useful if you edit the JSON files and want to reset them.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret" -OverwriteRequestBodies
```

### Example 6: Cleanup Old Reports

This example demonstrates how to remove older reports from the reports directory.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret" -CleanupOldReports
```

### Example 7: Use Device Authentication

This example shows how to use device authentication for Microsoft Graph API.

```powershell
.\Get-IntuneReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -UseDeviceAuthentication
```

These examples cover common use cases for the `Get-IntuneReport` script. Adjust the parameters as needed to fit your specific requirements.

## 🔢 Parameters

### -SavePath
```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 0
Default value: "$env:TEMP\IntuneReports"
Description: The path where the reports will be saved. Default is $env:TEMP\IntuneReports
```

### -FormatChoice

```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 1
Default value: csv
Description: The format of the report. Valid formats are csv and json. Default is csv
```

### -EndpointVersion

```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 2
Default value: beta
Description: The Microsoft Graph API version to use. Valid endpoints are v1.0 and beta. Default is beta
```

### -ModuleNames

```yaml
Type: Object
Parameter Sets: (All)
Required: False
Position: 3
Default value: Microsoft.Graph.Authentication
Description: Module Name to connect to Graph. Default is Microsoft.Graph.Authentication
```

### -PackageProvider

```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 4
Default value: NuGet
Description: If not specified, the default value NuGet is used for PackageProvider
```

### -TenantId

```yaml
Type: String
Parameter Sets: ClientSecret, ClientCertificateThumbprint, UseDeviceAuthentication, Interactive
Required: True
Position: 5
Description: Tenant Id or name to connect to
```

### -ClientId

```yaml
Type: String
Parameter Sets: ClientSecret, ClientCertificateThumbprint, UseDeviceAuthentication, Interactive
Required: True
Position: 6
Description: Client Id (App Registration) to connect to
```

### -ClientSecret

```yaml
Type: String
Parameter Sets: ClientSecret
Required: True
Position: 7
Description: Client secret for authentication
```

### -ClientCertificateThumbprint

```yaml
Type: String
Parameter Sets: ClientCertificateThumbprint
Required: True
Position: 8
Description: Client certificate thumbprint for authentication
```

### -UseDeviceAuthentication

```yaml
Type: Switch
Parameter Sets: UseDeviceAuthentication
Required: True
Position: 9
Description: Use device authentication for Microsoft Graph API
```

### -RequiredScopes

```yaml
Type: String[]
Parameter Sets: (All)
Required: False
Position: 10
Default value: Reports.Read.All
Description: The scopes required for Microsoft Graph API access. Default is Reports.Read.All
```

### -ModuleScope

```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 11
Default value: CurrentUser
Description: Specifies the scope for installing the module. Default is CurrentUser
```

### -OverwriteRequestBodies

```yaml
Type: Switch
Parameter Sets: (All)
Required: False
Position: 12
Description: Always overwrite existing JSON file
```

### -MaxRetries

```yaml
Type: Int
Parameter Sets: (All)
Required: False
Position: 13
Default value: 10
Description: Number of retries for polling report status. Default is 10
```

### -SecondsToWait

```yaml
Type: Int
Parameter Sets: (All)
Required: False
Position: 14
Default value: 5
Description: Seconds to wait between polling report status. Default is 5
```

### -CleanupOldReports

```yaml
Type: Switch
Parameter Sets: (All)
Required: False
Position: 15
Description: Cleanup old reports from the reports directory
```

### -ReportNames

```yaml
Type: String[]
Parameter Sets: (All)
Required: False
Position: 16
Description: List of report names to run
```

### -LogId

```yaml
Type: String
Parameter Sets: (All)
Required: False
Position: 17
Default value: $($MyInvocation.MyCommand).Name
Description: Component name for logging
```

### -ResetLog

```yaml
Type: Switch
Parameter Sets: (All)
Required: False
Position: 18
Description: ResetLog: Pass this parameter to reset the log file
```
