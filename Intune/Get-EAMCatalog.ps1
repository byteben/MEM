<#
.Synopsis
Created on:   26/09/2026
Created by:   Ben Whitmore
Filename:     Get-EAMCatalog.ps1

Get the Microsoft Enterprise App Management catalog using the Microsoft Graph PowerShell SDK.
The catalog APIs are only available on the Microsoft Graph beta endpoint.

.Description
This script gets the Microsoft Enterprise App Management catalog packages from Microsoft Graph, following @odata.nextLink to get every page.

Each catalog package includes the publisher, product, branch, architecture, locales, version and auto-update capability.
Pass -DetailedInfo to also get the win32CatalogApp template for each package using convertFromMobileAppCatalogPackage. The template includes the installer type, file name, install and uninstall commands, MSI information, detection rules and requirement rules.

The createdDateTime and lastModifiedDateTime values in the template are placeholders, so they cannot be used to tell when an app was updated in the catalog.
To track catalog updates, save the output on a schedule and compare each run by BranchId.

By default the catalog package records are returned. Pass -DetailedInfo to include the template data, or -HtmlReport to save and open an HTML report instead.
The installer type (MSI or EXE) is only available with -DetailedInfo because it is not included on the catalog package.
Records are sorted by Publisher, Product, Branch and CatalogVersion because Graph returns catalog packages in no particular order.

.NOTES
PowerShell 5.1 or later is required to run this script.
Microsoft.Graph.Authentication is required.
Required delegated scope: DeviceManagementApps.Read.All

---------------------------------------------------------------------------------
LEGAL DISCLAIMER

The PowerShell script provided is shared with the community as-is
The author and co-author(s) make no warranties or guarantees regarding its functionality, reliability, or suitability for any specific purpose
Please note that the script may need to be modified or adapted to fit your specific environment or requirements
It is recommended to thoroughly test the script in a non-production environment before using it in a live or critical system
The author and co-author(s) cannot be held responsible for any damages, losses, or adverse effects that may arise from the use of this script
You assume all risks and responsibilities associated with its usage
---------------------------------------------------------------------------------

.PARAMETER DetailedInfo
Also get the win32CatalogApp template for each package and add its description, installer type, file name, size, install and uninstall commands, MSI information, detection versions, detection rules, requirement rules and install behaviour to each record.
Templates are requested using Graph $batch, 20 packages per request, and throttled requests are retried.
Use -ProductName or -Publisher with -DetailedInfo so templates are only requested for matching packages.

.PARAMETER HtmlReport
Save an HTML report to the temp folder as EAM-Catalog.html and open it in the default browser. The file is overwritten on each run.
The report shows summary counts, a search box, filters and apps grouped by publisher and product. Pass -DetailedInfo to include the template data for each variant.
Records are not returned to the pipeline when -HtmlReport is passed.

.PARAMETER ProductName
Optional filter for productDisplayName. Matches anywhere in the product name, for example chrome matches Google Chrome. The filter is applied after the catalog is retrieved from Graph.

.PARAMETER Publisher
Optional filter for publisherDisplayName. Matches anywhere in the publisher name, for example google matches Google LLC. The filter is applied after the catalog is retrieved from Graph.

.PARAMETER PageSize
The number of records to request per Graph page. Defaults to 100.

.PARAMETER TenantId
The Tenant Id if using the Application authentication flow. Not required for Delegated authentication.

.PARAMETER ClientId
The Application (client) Id if using the Application authentication flow. Not required for Delegated authentication.

.PARAMETER CertThumbprint
The certificate thumbprint if using the Application authentication flow with certificate-based authentication. Not required for Delegated authentication.

.PARAMETER ClientSecret
The client secret if using the Application authentication flow with a client secret. Not required for Delegated authentication.

.PARAMETER AuthFlow
The authentication flow to use. Valid authentication flows are 'ApplicationCertificate', 'ApplicationClientSecret', and 'Delegated'. Default is 'Delegated'.

.EXAMPLE
.\Get-EAMCatalog.ps1

Return the basic record for every package in the catalog (publisher, product, version, branch, architecture, locales and auto-update capability).

.EXAMPLE
$catalog = .\Get-EAMCatalog.ps1

Capture the catalog in a variable so it can be queried without calling Graph again.

.EXAMPLE
.\Get-EAMCatalog.ps1 -ProductName chrome

Return packages with chrome anywhere in the product name.

.EXAMPLE
.\Get-EAMCatalog.ps1 -Publisher mozilla

Return packages with mozilla anywhere in the publisher name.

.EXAMPLE
.\Get-EAMCatalog.ps1 -ProductName firefox -Publisher mozilla

Return Mozilla packages with firefox anywhere in the product name.

.EXAMPLE
.\Get-EAMCatalog.ps1 -DetailedInfo -ProductName chrome

Return detailed records for Chrome packages. Filtering with -ProductName or -Publisher first means template data is only requested for matching packages, which is much faster than -DetailedInfo on its own.

.EXAMPLE
.\Get-EAMCatalog.ps1 -DetailedInfo

Return detailed records for every package, including installer type, command lines, MSI information, detection rules and requirement rules. This requests template data for the whole catalog so takes longer.

.EXAMPLE
.\Get-EAMCatalog.ps1 | Where-Object { $_.Architecture -match 'arm64' }

Return packages that support arm64.

.EXAMPLE
.\Get-EAMCatalog.ps1 | Where-Object { $_.AutoUpdateCapable }

Return packages that are auto-update capable.

.EXAMPLE
.\Get-EAMCatalog.ps1 | Group-Object Architecture | Select-Object Count, Name

Count packages by architecture.

.EXAMPLE
.\Get-EAMCatalog.ps1 -DetailedInfo | Where-Object { $_.InstallerType -eq 'MSI' }

Return MSI packages. InstallerType is only available with -DetailedInfo.

.EXAMPLE
.\Get-EAMCatalog.ps1 -DetailedInfo -Publisher google | Where-Object { $_.InstallerType -eq 'EXE' -and $_.AutoUpdateCapable } | Select-Object Product, Branch, CatalogVersion, InstallCommandLine

Return auto-update capable EXE packages from Google, showing the install command line.

.EXAMPLE
.\Get-EAMCatalog.ps1 -HtmlReport

Save and open an HTML report of the whole catalog with basic info. Includes summary cards, search, filters and apps grouped by publisher and product.

.EXAMPLE
.\Get-EAMCatalog.ps1 -HtmlReport -DetailedInfo

Save and open a full HTML report of the whole catalog with all details. Click a variant to see its description, installer type, file name, size, install and uninstall commands, MSI codes, detection rules, requirement rules and install behaviour.

.EXAMPLE
.\Get-EAMCatalog.ps1 -HtmlReport -DetailedInfo -ProductName chrome

Save and open a full HTML report with all details for Chrome packages only. Faster than a full catalog report because template data is only requested for matching packages.

.EXAMPLE
.\Get-EAMCatalog.ps1 -HtmlReport -DetailedInfo -Publisher mozilla

Save and open a full HTML report with all details for Mozilla packages only.

.EXAMPLE
.\Get-EAMCatalog.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertThumbprint "your-cert-thumbprint" -AuthFlow "ApplicationCertificate" -DetailedInfo

Return detailed records using an app registration with certificate authentication.

.EXAMPLE
.\Get-EAMCatalog.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret" -AuthFlow "ApplicationClientSecret" -HtmlReport

Save and open an HTML report using an app registration with client secret authentication.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 0, HelpMessage = 'Include win32CatalogApp template data for each package')]
    [switch]$DetailedInfo,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1, HelpMessage = 'Save an HTML report to the temp folder and open it instead of returning records')]
    [switch]$HtmlReport,


    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 2, HelpMessage = 'Optional filter for productDisplayName, matches anywhere in the name')]
    [string]$ProductName,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 3, HelpMessage = 'Optional filter for publisherDisplayName, matches anywhere in the name')]
    [string]$Publisher,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 4, HelpMessage = 'The number of records to request per Graph page')]
    [ValidateRange(1, 999)]
    [int]$PageSize = 100,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 5, HelpMessage = 'Tenant Id if using the Application auth flow')]
    [string]$TenantId = '',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 6, HelpMessage = 'App registration / client Id if using the Application auth flow')]
    [string]$ClientId = '',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 7, HelpMessage = 'Certificate thumbprint if using the Application auth flow with certificate authentication')]
    [string]$CertThumbprint = '',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 8, HelpMessage = 'Client Secret if using the Application auth flow with client secret')]
    [string]$ClientSecret = '',

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 9, HelpMessage = 'The authentication flow to use')]
    [ValidateSet('ApplicationCertificate', 'ApplicationClientSecret', 'Delegated')]
    [string]$AuthFlow = 'Delegated'
)

function Initialize-MgGraphModule {
    [CmdletBinding()]
    param ()

    $moduleName = 'Microsoft.Graph.Authentication'

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw ("PowerShell module '{0}' was not found. Install it with 'Install-Module {0} -Scope CurrentUser' and run this script again." -f $moduleName)
    }

    Import-Module -Name $moduleName -ErrorAction Stop
}

function Get-MgGraphPaged {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Graph URI to call')]
        [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    $page = 1

    while ($nextUri) {
        Write-Verbose ("Calling Graph page {0}: {1}" -f $page, $nextUri)
        Write-Progress -Activity 'Getting EAM catalog packages' -Status ("Page {0} ({1} packages so far)" -f $page, $items.Count)

        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop

        if ($response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }

        $nextUri = $response.'@odata.nextLink'
        $page++
    }

    Write-Progress -Activity 'Getting EAM catalog packages' -Completed

    return $items
}

function Get-EAMCatalogAppTemplateBatch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Catalog package Ids')]
        [string[]]$CatalogPackageIds,

        [Parameter(Mandatory = $false, HelpMessage = 'Maximum retry attempts for throttled requests')]
        [int]$MaxRetries = 5
    )

    $batchSize = 20
    $results = @{}
    $batchUri = 'https://graph.microsoft.com/beta/$batch'
    $totalBatches = [math]::Ceiling($CatalogPackageIds.Count / $batchSize)

    for ($index = 0; $index -lt $CatalogPackageIds.Count; $index += $batchSize) {
        $batchNumber = [math]::Floor($index / $batchSize) + 1
        $pending = @($CatalogPackageIds[$index..([math]::Min($index + $batchSize, $CatalogPackageIds.Count) - 1)])
        $attempt = 0

        Write-Progress -Activity 'Getting EAM catalog template data' -Status ("Batch {0} of {1}" -f $batchNumber, $totalBatches) -PercentComplete (($batchNumber / [math]::Max($totalBatches, 1)) * 100)

        while ($pending.Count -gt 0) {
            $attempt++

            $requests = for ($requestIndex = 0; $requestIndex -lt $pending.Count; $requestIndex++) {
                $catalogPackageId = $pending[$requestIndex]
                @{
                    id     = [string]$requestIndex
                    method = 'GET'
                    url    = "/deviceAppManagement/mobileApps/convertFromMobileAppCatalogPackage(mobileAppCatalogPackageId='$catalogPackageId')"
                }
            }

            $body = @{ requests = @($requests) } | ConvertTo-Json -Depth 5
            $response = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json' -ErrorAction Stop

            $throttled = [System.Collections.Generic.List[string]]::new()
            $retryAfter = 0

            foreach ($item in $response.responses) {
                $catalogPackageId = $pending[[int]$item.id]

                if ([int]$item.status -eq 200) {
                    $results[$catalogPackageId] = [pscustomobject]@{ Template = $item.body; Error = $null }
                }
                elseif (([int]$item.status -eq 429 -or [int]$item.status -ge 500) -and $attempt -le $MaxRetries) {
                    $throttled.Add($catalogPackageId)
                    $itemRetryAfter = 10
                    if ($item.headers -and $item.headers.'Retry-After') {
                        $itemRetryAfter = [int]$item.headers.'Retry-After'
                    }
                    $retryAfter = [math]::Max($retryAfter, $itemRetryAfter)
                }
                else {
                    $errorMessage = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
                    $results[$catalogPackageId] = [pscustomobject]@{ Template = $null; Error = $errorMessage }
                }
            }

            $pending = @($throttled)

            if ($pending.Count -gt 0) {
                Write-Warning ("Graph throttled {0} request(s) in batch {1}. Retrying in {2} second(s)" -f $pending.Count, $batchNumber, $retryAfter)
                Start-Sleep -Seconds $retryAfter
            }
        }
    }

    Write-Progress -Activity 'Getting EAM catalog template data' -Completed

    return $results
}

function Get-EAMInstallerType {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Generated win32CatalogApp template')]
        [object]$Template
    )

    if ($Template.msiInformation) {
        return 'MSI'
    }
    elseif ($Template.setupFilePath -match '\.exe$') {
        return 'EXE'
    }
    elseif ($Template.setupFilePath -match '\.msi$') {
        return 'MSI'
    }
    else {
        return 'Unknown'
    }
}

function Get-EAMDetectionVersions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Generated win32CatalogApp template')]
        [object]$Template
    )

    $versions = @(
        $Template.rules |
            Where-Object { $_.ruleType -eq 'detection' } |
            ForEach-Object {
                if ($_.productCode -and $_.productVersion) {
                    $_.productVersion
                }
                elseif (($_.operationType -in @('version', 'appVersion') -or $_.valueName -eq 'DisplayVersion') -and $_.comparisonValue) {
                    $_.comparisonValue
                }
            }
    ) | Sort-Object -Unique

    return $versions
}

function Format-EAMRule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Win32 app detection or requirement rule')]
        [object]$Rule
    )

    $bitness = if ($Rule.check32BitOn64System -or $Rule.runAs32Bit) { ' (32-bit)' } else { '' }

    switch -Wildcard ($Rule.'@odata.type') {
        '*ProductCodeRule' {
            $version = if ($Rule.productVersion) { " productVersion {0} {1}" -f $Rule.productVersionOperator, $Rule.productVersion } else { '' }
            return ("MSI: {0}{1}" -f $Rule.productCode, $version)
        }
        '*FileSystemRule' {
            return ("File: {0}\{1} {2} {3} {4}{5}" -f $Rule.path, $Rule.fileOrFolderName, $Rule.operationType, $Rule.operator, $Rule.comparisonValue, $bitness).Trim()
        }
        '*RegistryRule' {
            return ("Registry: {0}\{1} {2} {3} {4}{5}" -f $Rule.keyPath, $Rule.valueName, $Rule.operationType, $Rule.operator, $Rule.comparisonValue, $bitness).Trim()
        }
        '*PowerShellScriptRule' {
            return ("Script: {0} {1} {2} {3} (runAs {4}){5}" -f $Rule.displayName, $Rule.operationType, $Rule.operator, $Rule.comparisonValue, $Rule.runAsAccount, $bitness)
        }
        default {
            return ("{0}: {1}" -f $Rule.'@odata.type', ($Rule | ConvertTo-Json -Compress -Depth 3))
        }
    }
}

function Get-EAMRuleText {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Generated win32CatalogApp template')]
        [object]$Template,

        [Parameter(Mandatory = $true, HelpMessage = 'Rule type to return')]
        [ValidateSet('detection', 'requirement')]
        [string]$RuleType
    )

    $ruleText = @(
        $Template.rules |
            Where-Object { $_.ruleType -eq $RuleType } |
            ForEach-Object { Format-EAMRule -Rule $_ }
    )

    return ($ruleText -join [Environment]::NewLine)
}

function ConvertTo-EAMCatalogRecord {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Catalog package object')]
        [object]$Package,

        [Parameter(Mandatory = $false, HelpMessage = 'Generated win32CatalogApp template')]
        [object]$Template,

        [Parameter(Mandatory = $false, HelpMessage = 'Error returned when getting the template')]
        [string]$TemplateError
    )

    $record = [ordered]@{
        Publisher                  = $Package.publisherDisplayName
        Product                    = $Package.productDisplayName
        CatalogVersion             = $Package.versionDisplayName
        Branch                     = $Package.branchDisplayName
        Architecture               = $Package.applicableArchitectures
        Locales                    = ($Package.locales -join ',')
        AutoUpdateCapable          = $Package.packageAutoUpdateCapable
        ProductId                  = $Package.productId
        BranchId                   = $Package.branchId
        CatalogPackageId           = $Package.id
    }

    if ($Template) {
        $detectionVersions = Get-EAMDetectionVersions -Template $Template

        $record['Description'] = $Template.description
        $record['InstallerType'] = Get-EAMInstallerType -Template $Template
        $record['AllowedArchitectures'] = $Template.allowedArchitectures
        $record['FileName'] = $Template.fileName
        $record['SetupFilePath'] = $Template.setupFilePath
        $record['Size'] = $Template.size
        $record['InstallCommandLine'] = $Template.installCommandLine
        $record['UninstallCommandLine'] = $Template.uninstallCommandLine
        $record['DisplayVersion'] = $Template.displayVersion
        $record['MsiProductVersion'] = $Template.msiInformation.productVersion
        $record['MsiProductCode'] = $Template.msiInformation.productCode
        $record['MsiUpgradeCode'] = $Template.msiInformation.upgradeCode
        $record['DetectionVersions'] = ($detectionVersions -join ',')
        $record['DetectionRules'] = Get-EAMRuleText -Template $Template -RuleType 'detection'
        $record['RequirementRules'] = Get-EAMRuleText -Template $Template -RuleType 'requirement'
        $record['MinimumWindowsRelease'] = $Template.minimumSupportedWindowsRelease
        $record['RunAsAccount'] = $Template.installExperience.runAsAccount
        $record['MaxRunTimeInMinutes'] = $Template.installExperience.maxRunTimeInMinutes
        $record['RestartBehavior'] = $Template.installExperience.deviceRestartBehavior
        $record['MobileAppCatalogPackageId'] = $Template.mobileAppCatalogPackageId
        $record['TemplateError'] = $null
    }
    elseif ($TemplateError) {
        $record['TemplateError'] = $TemplateError
    }

    return [pscustomobject]$record
}

function Get-EAMCatalogReportSummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'EAM catalog records')]
        [AllowEmptyCollection()]
        [object[]]$Records
    )

    $hasInstallerType = @($Records | Where-Object { $_.PSObject.Properties.Name -contains 'InstallerType' }).Count -gt 0

    $summary = [ordered]@{
        Publishers        = @($Records | ForEach-Object { $_.Publisher } | Sort-Object -Unique).Count
        Products          = @($Records | ForEach-Object { if ($_.ProductId) { $_.ProductId } else { $_.Product } } | Sort-Object -Unique).Count
        Variants          = @($Records | ForEach-Object { if ($_.BranchId) { $_.BranchId } else { $_.CatalogPackageId } } | Sort-Object -Unique).Count
        Packages          = $Records.Count
        AutoUpdateCapable = @($Records | Where-Object { $_.AutoUpdateCapable -eq $true }).Count
        x64               = @($Records | Where-Object { ([string]$_.Architecture -split ',') -contains 'x64' }).Count
        x86               = @($Records | Where-Object { ([string]$_.Architecture -split ',') -contains 'x86' }).Count
        arm64             = @($Records | Where-Object { ([string]$_.Architecture -split ',') -contains 'arm64' }).Count
        MSI               = $null
        EXE               = $null
    }

    if ($hasInstallerType) {
        $summary['MSI'] = @($Records | Where-Object { $_.InstallerType -eq 'MSI' }).Count
        $summary['EXE'] = @($Records | Where-Object { $_.InstallerType -eq 'EXE' }).Count
    }

    return [pscustomobject]$summary
}

function New-EAMCatalogHtmlReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'EAM catalog records')]
        [AllowEmptyCollection()]
        [object[]]$Records,

        [Parameter(Mandatory = $true, HelpMessage = 'Path to save the HTML report to')]
        [string]$Path,

        [Parameter(Mandatory = $false, HelpMessage = 'Product name filter used to get the records')]
        [string]$ProductName,

        [Parameter(Mandatory = $false, HelpMessage = 'Publisher filter used to get the records')]
        [string]$Publisher,

        [Parameter(Mandatory = $false, HelpMessage = 'Records include template data')]
        [switch]$DetailedInfo
    )

    $meta = [ordered]@{
        Generated   = (Get-Date).ToString('o')
        ProductName = $ProductName
        Publisher   = $Publisher
        Detailed    = [bool]$DetailedInfo
        Summary     = Get-EAMCatalogReportSummary -Records $Records
    }

    $dataJson = if ($Records.Count -gt 0) { ConvertTo-Json -InputObject @($Records) -Depth 3 -Compress } else { '[]' }
    $metaJson = ConvertTo-Json -InputObject $meta -Depth 3 -Compress

    $dataJson = $dataJson.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
    $metaJson = $metaJson.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')

    $template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Microsoft Enterprise App Management Catalog</title>
<style>
:root {
  --bg: #f6f7f9; --panel: #ffffff; --text: #1f2328; --muted: #656d76; --border: #e3e6ea;
  --accent: #4f46e5; --accent-soft: #eef2ff; --ok: #15803d; --ok-soft: #dcfce7;
  --hover: #f8fafc; --code: #f3f4f6; --shadow: 0 1px 2px rgba(16, 24, 40, .06), 0 1px 3px rgba(16, 24, 40, .08);
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #0d1117; --panel: #161b22; --text: #e6edf3; --muted: #8b949e; --border: #30363d;
    --accent: #a5b4fc; --accent-soft: #1e1b4b; --ok: #4ade80; --ok-soft: #14532d;
    --hover: #1c2129; --code: #0d1117; --shadow: none;
  }
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--text); font: 14px/1.5 "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, sans-serif; }
.wrap { max-width: 1280px; margin: 0 auto; padding: 32px 24px 64px; }
header h1 { margin: 0 0 4px; font-size: 26px; font-weight: 650; letter-spacing: -.01em; }
.sub { color: var(--muted); display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
.chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 10px; border-radius: 999px; background: var(--accent-soft); color: var(--accent); font-size: 12px; font-weight: 600; }
.chip.muted { background: var(--code); color: var(--muted); }
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; margin: 24px 0; }
.card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; box-shadow: var(--shadow); }
.card .value { font-size: 26px; font-weight: 650; line-height: 1.2; }
.card .label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
.toolbar { position: sticky; top: 0; z-index: 5; display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding: 12px 0; background: var(--bg); }
.toolbar input, .toolbar select, .toolbar button { font: inherit; color: var(--text); background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 7px 10px; }
.toolbar input { flex: 1 1 260px; min-width: 200px; }
.toolbar input:focus, .toolbar select:focus { outline: 2px solid var(--accent); outline-offset: -1px; }
.toolbar button { cursor: pointer; }
.toolbar button:hover { background: var(--hover); }
.count { color: var(--muted); margin-left: auto; font-size: 13px; }
details.pub { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; margin-bottom: 10px; box-shadow: var(--shadow); overflow: hidden; }
details > summary { list-style: none; cursor: pointer; display: flex; align-items: center; gap: 10px; }
details > summary::-webkit-details-marker { display: none; }
details > summary::before { content: ""; width: 7px; height: 7px; border-right: 2px solid var(--muted); border-bottom: 2px solid var(--muted); transform: rotate(-45deg); transition: transform .15s; flex: none; }
details[open] > summary::before { transform: rotate(45deg); }
details.pub > summary { padding: 14px 18px; font-weight: 600; font-size: 15px; }
details.pub > summary:hover, details.prod > summary:hover { background: var(--hover); }
details.prod { border-top: 1px solid var(--border); }
details.prod > summary { padding: 10px 18px 10px 38px; font-weight: 550; }
.meta { color: var(--muted); font-weight: 400; font-size: 13px; margin-left: auto; display: flex; gap: 8px; align-items: center; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { text-align: left; color: var(--muted); font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; padding: 8px 12px; border-top: 1px solid var(--border); background: var(--hover); }
td { padding: 9px 12px; border-top: 1px solid var(--border); vertical-align: top; }
th:first-child, td:first-child { padding-left: 58px; }
tr.row { cursor: pointer; }
tr.row:hover td { background: var(--hover); }
tr.row.open td { background: var(--accent-soft); }
.pill { display: inline-block; padding: 1px 8px; border-radius: 999px; background: var(--code); color: var(--muted); font-size: 12px; font-weight: 600; margin-right: 4px; }
.pill.ok { background: var(--ok-soft); color: var(--ok); }
.pill.accent { background: var(--accent-soft); color: var(--accent); }
.version { font-family: "Cascadia Mono", Consolas, ui-monospace, monospace; font-size: 12.5px; }
td.detail { padding: 16px 18px 20px 58px; background: var(--panel); }
dl { display: grid; grid-template-columns: 200px 1fr; gap: 8px 16px; margin: 0; }
dt { color: var(--muted); font-weight: 600; font-size: 12px; }
dd { margin: 0; min-width: 0; word-break: break-word; }
pre, code { font-family: "Cascadia Mono", Consolas, ui-monospace, monospace; font-size: 12px; background: var(--code); border: 1px solid var(--border); border-radius: 6px; }
pre { margin: 0; padding: 8px 10px; white-space: pre-wrap; }
code { padding: 2px 6px; display: inline-block; }
.empty { text-align: center; color: var(--muted); padding: 48px; background: var(--panel); border: 1px dashed var(--border); border-radius: 12px; }
footer { color: var(--muted); font-size: 12px; margin-top: 24px; text-align: center; }
@media (max-width: 720px) { dl { grid-template-columns: 1fr; } th:first-child, td:first-child, td.detail { padding-left: 16px; } .hide-sm { display: none; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Microsoft Enterprise App Management Catalog</h1>
    <div class="sub" data-id="sub"></div>
  </header>
  <section class="cards" data-id="cards"></section>
  <div class="toolbar">
    <input type="search" data-id="search" placeholder="Search publisher, product, variant or version" aria-label="Search">
    <select data-id="arch" aria-label="Architecture" hidden><option value="">All architectures</option></select>
    <select data-id="auto" aria-label="Auto-update" hidden><option value="">Auto-update: any</option></select>
    <select data-id="inst" aria-label="Installer type" hidden><option value="">All installers</option></select>
    <button type="button" data-id="expand">Expand all</button>
    <button type="button" data-id="collapse">Collapse all</button>
    <span class="count" data-id="count"></span>
  </div>
  <main data-id="list"></main>
  <footer>Generated by Get-EAMCatalog.ps1 using Microsoft Graph beta</footer>
</div>
<script type="application/json" data-id="eam-data">__EAM_DATA__</script>
<script type="application/json" data-id="eam-meta">__EAM_META__</script>
<script>
(function () {
  var $ = function (id) { return document.querySelector('[data-id="' + id + '"]'); };
  var data = JSON.parse($('eam-data').textContent);
  var meta = JSON.parse($('eam-meta').textContent);
  var detailed = meta.Detailed === true;

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (text !== undefined && text !== null) { e.textContent = String(text); }
    return e;
  }

  function archList(r) { return String(r.Architecture || '').split(',').filter(Boolean); }

  var sub = $('sub');
  sub.appendChild(el('span', null, 'Generated ' + new Date(meta.Generated).toLocaleString()));
  sub.appendChild(el('span', 'chip', detailed ? 'Detailed info' : 'Basic info'));
  if (meta.ProductName) { sub.appendChild(el('span', 'chip muted', 'Product: ' + meta.ProductName)); }
  if (meta.Publisher) { sub.appendChild(el('span', 'chip muted', 'Publisher: ' + meta.Publisher)); }

  var s = meta.Summary || {};
  var cards = [['Publishers', s.Publishers], ['Products', s.Products], ['Variants', s.Variants], ['Auto-update capable', s.AutoUpdateCapable], ['x64', s.x64], ['x86', s.x86], ['arm64', s.arm64]];
  if (detailed) { cards.push(['MSI', s.MSI], ['EXE', s.EXE]); }
  cards.forEach(function (c) {
    var card = el('div', 'card');
    card.appendChild(el('div', 'value', (c[1] || 0).toLocaleString()));
    card.appendChild(el('div', 'label', c[0]));
    $('cards').appendChild(card);
  });

  function addFilterOptions(id, options) {
    var select = $(id);
    var valid = options.filter(function (o) { return o[2] > 0; });
    valid.forEach(function (o) {
      var option = el('option', null, o[1] + ' (' + o[2].toLocaleString() + ')');
      option.value = o[0];
      select.appendChild(option);
    });
    select.hidden = valid.length < 2;
  }

  function countWhere(test) { return data.filter(test).length; }

  var archValues = [];
  data.forEach(function (r) { archList(r).forEach(function (x) { if (archValues.indexOf(x) === -1) { archValues.push(x); } }); });
  addFilterOptions('arch', archValues.sort().map(function (x) { return [x, x, countWhere(function (r) { return archList(r).indexOf(x) !== -1; })]; }));
  addFilterOptions('auto', [
    ['yes', 'Auto-update capable', countWhere(function (r) { return r.AutoUpdateCapable === true; })],
    ['no', 'Not auto-update capable', countWhere(function (r) { return r.AutoUpdateCapable !== true; })]
  ]);
  if (detailed) {
    var installerValues = [];
    data.forEach(function (r) { if (r.InstallerType && installerValues.indexOf(r.InstallerType) === -1) { installerValues.push(r.InstallerType); } });
    addFilterOptions('inst', installerValues.sort().map(function (x) { return [x, x, countWhere(function (r) { return r.InstallerType === x; })]; }));
  }

  var groups = [];
  var pubIndex = {};
  data.forEach(function (r) {
    var pubName = r.Publisher || 'Unknown publisher';
    var pub = pubIndex[pubName];
    if (!pub) { pub = { name: pubName, products: [], index: {} }; pubIndex[pubName] = pub; groups.push(pub); }
    var key = r.ProductId || r.Product;
    var prod = pub.index[key];
    if (!prod) { prod = { name: r.Product || 'Unknown product', rows: [] }; pub.index[key] = prod; pub.products.push(prod); }
    prod.rows.push(r);
  });

  function formatSize(bytes) {
    var n = Number(bytes);
    if (!n) { return ''; }
    return n >= 1073741824 ? (n / 1073741824).toFixed(2) + ' GB' : (n / 1048576).toFixed(1) + ' MB';
  }

  var fields = [
    ['Description', 'Description'], ['Installer type', 'InstallerType'], ['File name', 'FileName'], ['Size', 'Size', 'size'],
    ['Install command', 'InstallCommandLine', 'code'], ['Uninstall command', 'UninstallCommandLine', 'code'],
    ['Display version', 'DisplayVersion'], ['MSI product code', 'MsiProductCode', 'code'], ['MSI upgrade code', 'MsiUpgradeCode', 'code'],
    ['Detection versions', 'DetectionVersions'], ['Detection rules', 'DetectionRules', 'pre'], ['Requirement rules', 'RequirementRules', 'pre', 'None'],
    ['Allowed architectures', 'AllowedArchitectures'], ['Minimum Windows release', 'MinimumWindowsRelease'], ['Run as', 'RunAsAccount'],
    ['Max run time (minutes)', 'MaxRunTimeInMinutes'], ['Restart behaviour', 'RestartBehavior'], ['Template error', 'TemplateError'],
    ['Product ID', 'ProductId', 'code'], ['Branch ID', 'BranchId', 'code'], ['Catalog package ID', 'CatalogPackageId', 'code']
  ];

  function detailPanel(r) {
    var dl = el('dl');
    fields.forEach(function (f) {
      if (!(f[1] in r) && !f[3]) { return; }
      if (f[3] && !detailed) { return; }
      var value = r[f[1]];
      if (f[2] === 'size') { value = formatSize(value); }
      if (value === null || value === undefined || value === '') {
        if (!f[3]) { return; }
        value = f[3];
      }
      dl.appendChild(el('dt', null, f[0]));
      var dd = el('dd');
      if (f[2] === 'pre' && value !== f[3]) { dd.appendChild(el('pre', null, value)); }
      else if (f[2] === 'code') { dd.appendChild(el('code', null, value)); }
      else { dd.textContent = String(value); }
      dl.appendChild(dd);
    });
    return dl;
  }

  var columns = detailed ? 6 : 5;

  function variantRow(r, tbody) {
    var tr = el('tr', 'row');
    tr.appendChild(el('td', null, r.Branch));
    var v = el('td'); v.appendChild(el('span', 'version', r.CatalogVersion)); tr.appendChild(v);
    var a = el('td'); archList(r).forEach(function (x) { a.appendChild(el('span', 'pill', x)); }); tr.appendChild(a);
    tr.appendChild(el('td', 'hide-sm', r.Locales));
    var u = el('td'); u.appendChild(el('span', r.AutoUpdateCapable === true ? 'pill ok' : 'pill', r.AutoUpdateCapable === true ? 'Yes' : 'No')); tr.appendChild(u);
    if (detailed) { var i = el('td'); if (r.InstallerType) { i.appendChild(el('span', 'pill accent', r.InstallerType)); } tr.appendChild(i); }
    var detailRow = null;
    tr.addEventListener('click', function () {
      if (detailRow) { detailRow.remove(); detailRow = null; tr.classList.remove('open'); return; }
      detailRow = el('tr');
      var td = el('td', 'detail'); td.colSpan = columns; td.appendChild(detailPanel(r));
      detailRow.appendChild(td);
      tr.after(detailRow);
      tr.classList.add('open');
    });
    tbody.appendChild(tr);
  }

  function productTable(rows) {
    var table = el('table');
    var head = el('tr');
    ['Variant', 'Version', 'Architecture', 'Locales', 'Auto-update'].concat(detailed ? ['Installer'] : []).forEach(function (h) {
      head.appendChild(el('th', h === 'Locales' ? 'hide-sm' : null, h));
    });
    var thead = el('thead'); thead.appendChild(head); table.appendChild(thead);
    var tbody = el('tbody');
    rows.forEach(function (r) { variantRow(r, tbody); });
    table.appendChild(tbody);
    return table;
  }

  function matches(r, q, arch, auto, inst) {
    if (q && [r.Publisher, r.Product, r.Branch, r.CatalogVersion].join(' ').toLowerCase().indexOf(q) === -1) { return false; }
    if (arch && archList(r).indexOf(arch) === -1) { return false; }
    if (auto === 'yes' && r.AutoUpdateCapable !== true) { return false; }
    if (auto === 'no' && r.AutoUpdateCapable === true) { return false; }
    if (inst && r.InstallerType !== inst) { return false; }
    return true;
  }

  function plural(n, word) { return n.toLocaleString() + ' ' + word + (n === 1 ? '' : 's'); }

  function render() {
    var q = $('search').value.trim().toLowerCase();
    var arch = $('arch').value, auto = $('auto').value, inst = $('inst').value;
    var active = !!(q || arch || auto || inst);
    var list = $('list');
    list.replaceChildren();
    var shown = 0;

    groups.forEach(function (pub) {
      var products = pub.products.map(function (p) {
        return { name: p.name, rows: p.rows.filter(function (r) { return matches(r, q, arch, auto, inst); }) };
      }).filter(function (p) { return p.rows.length > 0; });
      if (!products.length) { return; }

      var variants = products.reduce(function (n, p) { return n + p.rows.length; }, 0);
      shown += variants;

      var pd = el('details', 'pub');
      pd.open = active;
      var ps = el('summary');
      ps.appendChild(el('span', null, pub.name));
      var pm = el('span', 'meta');
      pm.appendChild(el('span', null, plural(products.length, 'product') + ' \u00b7 ' + plural(variants, 'variant')));
      ps.appendChild(pm);
      pd.appendChild(ps);

      products.forEach(function (p) {
        var d = el('details', 'prod');
        d.open = active;
        var sm = el('summary');
        sm.appendChild(el('span', null, p.name));
        var m = el('span', 'meta');
        var versions = p.rows.map(function (r) { return r.CatalogVersion; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; });
        versions.slice(0, 3).forEach(function (v) { m.appendChild(el('span', 'pill version', v)); });
        if (versions.length > 3) { m.appendChild(el('span', 'pill', '+' + (versions.length - 3))); }
        m.appendChild(el('span', null, plural(p.rows.length, 'variant')));
        sm.appendChild(m);
        d.appendChild(sm);
        var built = false;
        var build = function () { if (!built && d.open) { d.appendChild(productTable(p.rows)); built = true; } };
        d.addEventListener('toggle', build);
        build();
        pd.appendChild(d);
      });

      list.appendChild(pd);
    });

    if (!shown) { list.appendChild(el('div', 'empty', data.length ? 'No apps match your search or filters.' : 'No catalog packages were returned.')); }
    $('count').textContent = 'Showing ' + shown.toLocaleString() + ' of ' + plural(data.length, 'variant');
  }

  var timer = null;
  $('search').addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(render, 150); });
  ['arch', 'auto', 'inst'].forEach(function (id) { $(id).addEventListener('change', render); });
  $('expand').addEventListener('click', function () { document.querySelectorAll('details').forEach(function (d) { d.open = true; }); });
  $('collapse').addEventListener('click', function () { document.querySelectorAll('details').forEach(function (d) { d.open = false; }); });

  render();
})();
</script>
</body>
</html>
'@

    $html = $template.Replace('__EAM_DATA__', $dataJson).Replace('__EAM_META__', $metaJson)

    [System.IO.File]::WriteAllText($Path, $html, (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false))

    return $Path
}

function Invoke-EAMCatalog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 0)]
        [switch]$DetailedInfo,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1)]
        [switch]$HtmlReport,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 2)]
        [string]$ProductName,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 3)]
        [string]$Publisher,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 4)]
        [ValidateRange(1, 999)]
        [int]$PageSize = 100,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 5)]
        [string]$TenantId = '',

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 6)]
        [string]$ClientId = '',

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 7)]
        [string]$CertThumbprint = '',

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 8)]
        [string]$ClientSecret = '',

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 9)]
        [ValidateSet('ApplicationCertificate', 'ApplicationClientSecret', 'Delegated')]
        [string]$AuthFlow = 'Delegated'
    )

    Initialize-MgGraphModule

    $RequiredScopes = @('DeviceManagementApps.Read.All')

    Write-Host 'Connecting to Microsoft Graph...'

    switch ($AuthFlow) {

        'ApplicationCertificate' {
            Write-Host 'Using application authentication flow with certificate.'

            if ([string]::IsNullOrWhiteSpace($CertThumbprint) -or [string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId)) {
                Write-Error 'Some authentication information is missing. Please specify the CertThumbprint, ClientId and TenantId parameters.'
                exit
            }

            $stores = @('CurrentUser', 'LocalMachine')

            foreach ($store in $stores) {
                $certThumbprintFound = Get-ChildItem -Path "Cert:\$store\My" | Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -ExpandProperty Thumbprint

                if ($certThumbprintFound) {
                    Write-Host ("Certificate with thumbprint '{0}' found in the '{1}' store." -f $CertThumbprint, $store)
                    break
                }
            }

            if (-not $certThumbprintFound) {
                Write-Error ("No certificate found. Please ensure you have a certificate with the thumbprint '{0}' in either the Computer or User personal certificate store." -f $CertThumbprint)
                exit
            }

            try {
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to connect to Microsoft Graph. Error: $_"
                exit
            }
        }

        'ApplicationClientSecret' {
            Write-Host 'Using application authentication flow with client secret.'

            if ([string]::IsNullOrWhiteSpace($ClientSecret) -or [string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId)) {
                Write-Error 'Some authentication information is missing. Please specify the ClientSecret, ClientId and TenantId parameters.'
                exit
            }

            try {
                $clientSecretSecure = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
                $clientCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $clientSecretSecure
                Connect-MgGraph -ClientSecretCredential $clientCredential -TenantId $TenantId -NoWelcome -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to connect to Microsoft Graph. Error: $_"
                exit
            }
        }

        'Delegated' {
            Write-Host 'Using delegated authentication flow...'

            try {
                Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to connect to Microsoft Graph. Error: $_"
                exit
            }
        }

        default {
            Write-Error 'Invalid authentication flow. Please specify a valid authentication flow.'
            exit
        }
    }

    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppCatalogPackages?`$top=$PageSize"

    Write-Host 'Getting Enterprise App Management catalog packages' -ForegroundColor Cyan

    try {
        $packages = Get-MgGraphPaged -Uri $uri
    }
    catch {
        throw ("Failed to get Enterprise App Management catalog packages from Microsoft Graph. Error: {0}" -f $_)
    }

    Write-Host ("Retrieved {0} catalog package(s)" -f $packages.Count) -ForegroundColor Green

    if ($ProductName) {
        $packages = @($packages | Where-Object { $_.productDisplayName -like "*$ProductName*" })
    }

    if ($Publisher) {
        $packages = @($packages | Where-Object { $_.publisherDisplayName -like "*$Publisher*" })
    }

    if ($ProductName -or $Publisher) {
        $filters = @()
        if ($ProductName) { $filters += "product '$ProductName'" }
        if ($Publisher) { $filters += "publisher '$Publisher'" }
        Write-Host ("Found {0} catalog package(s) matching {1}" -f $packages.Count, ($filters -join ' and ')) -ForegroundColor Green
    }

    $records = [System.Collections.Generic.List[object]]::new()

    if ($DetailedInfo -and $packages.Count -gt 0) {
        Write-Host 'Getting template data using Graph batch requests' -ForegroundColor Cyan
        $templateResults = Get-EAMCatalogAppTemplateBatch -CatalogPackageIds @($packages | ForEach-Object { $_.id })
    }

    foreach ($package in $packages) {
        if ($DetailedInfo) {
            $templateResult = $templateResults[$package.id]

            if ($templateResult -and $templateResult.Template) {
                $records.Add((ConvertTo-EAMCatalogRecord -Package $package -Template $templateResult.Template))
            }
            else {
                $templateError = if ($templateResult) { $templateResult.Error } else { 'No template response returned' }
                Write-Warning ("Failed to get template data for '{0}' {1} ({2}): {3}" -f $package.productDisplayName, $package.versionDisplayName, $package.id, $templateError)
                $records.Add((ConvertTo-EAMCatalogRecord -Package $package -TemplateError $templateError))
            }
        }
        else {
            $records.Add((ConvertTo-EAMCatalogRecord -Package $package))
        }
    }

    $records = @($records | Sort-Object Publisher, Product, Branch, CatalogVersion)

    if ($HtmlReport) {
        $reportPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'EAM-Catalog.html'
        New-EAMCatalogHtmlReport -Records $records -Path $reportPath -ProductName $ProductName -Publisher $Publisher -DetailedInfo:$DetailedInfo | Out-Null
        Write-Host ("HTML report saved to {0}" -f $reportPath) -ForegroundColor Green
        Start-Process -FilePath $reportPath
        return
    }

    return $records
}

Invoke-EAMCatalog @PSBoundParameters
