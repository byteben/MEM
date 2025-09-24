<#
.SYNOPSIS
    Enumerates browser extensions for Edge, Chrome, and Firefox for all user profiles on the system.

.DESCRIPTION
    Scans user profile directories for browser extension configuration files, parses extension metadata,
    evaluates risk based on permissions, and outputs a summary table.

    ################# IMPORTANT ################# 
    The Risk categories are based on common extension permissions and their potential impact. 
    The list should be used as a guide only and not a definitive security assessment.
    Please review the permissions and functionality of each extension individually to determine its suitability for your environment.

.NOTES
    Author: Ben Whitmore@PatchMyPC
    Created: 22nd September 2025

    ################# DISCLAIMER #################
    Patch My PC provides scripts, macro, and other code examples for illustration only, without warranty 
    either expressed or implied, including but not limited to the implied warranties of merchantability 
    and/or fitness for a particular purpose. This script is provided 'AS IS' and Patch My PC does not 
    guarantee that the following script, macro, or code can or should be used in any situation or that 
    operation of the code will be error-free.

.PARAMETER BaseUserPath
Base path to user profiles. Default is "C:\Users\*\AppData".

.PARAMETER Browser
Filter results by browser type. Accepts "Edge", "Chrome", or "Firefox".

.PARAMETER RiskLevel        
Filter results by risk level. Accepts "Critical", "High", "Medium", "Low", or "Unknown".

.PARAMETER ExcludeBuiltIn
Exclude built-in extensions from results. This is set to true by default.

.PARAMETER IncludeBuiltInOnly
Include only built-in extensions in results.

.PARAMETER UserProfile
Filter results by specific user profile paths.

.PARAMETER ExportPath
Path to export results as a CSV file.

.PARAMETER Verbose
Use -Verbose to display detailed troubleshooting output.

.PARAMETER OutputFormat
Format of the output display. Accepts "Table" or "List". Default is "Table".

.EXAMPLE
    .\Get-BrowserExtensions.ps1 -Browser Chrome -RiskLevel High -ExcludeBuiltIn

.EXAMPLE
    .\Get-BrowserExtensions.ps1 -Verbose
    Display all results and print detailed troubleshooting information.
#>

[CmdletBinding()]
param(
    [string]$BaseUserPath = "C:\Users\*\AppData",
    [ValidateSet("Edge", "Chrome", "Firefox")]
    [string[]]$Browser,
    [ValidateSet("Critical", "High", "Medium", "Low", "Unknown")]
    [string[]]$RiskLevel,
    [switch]$ExcludeBuiltIn,
    [switch]$IncludeBuiltInOnly,
    [string[]]$UserProfile,
    [string]$ExportPath,
    [ValidateSet("Table", "List")]
    [string]$OutputFormat = "Table"
)

$browserPaths = @(
    @{ Pattern = "Local\Microsoft\Edge\User Data\*\Preferences"; Browser = "Edge"; Regex = '^(Default|Profile ?\d+)$' }
    @{ Pattern = "Local\Google\Chrome\User Data\*\Preferences"; Browser = "Chrome"; Regex = '^(Default|Profile ?\d+)$' }
    @{ Pattern = "Roaming\Mozilla\Firefox\Profiles\*\extensions.json"; Browser = "Firefox"; Regex = $null }
)

$apiRiskCategories = @{
    "Critical" = @("webRequest", "webRequestBlocking", "webRequestAuthProvider", "cookies", "nativeMessaging", "debugger", "scripting", "tabCapture", "management", "clipboardRead", "geolocation", "camera", "microphone", "desktopCapture")
    "High"     = @("history", "downloads", "devtools", "clipboardWrite", "contentSettings", "proxy", "webNavigation", "privacy", "tabs", "identity", "pageCapture")
    "Medium"   = @("bookmarks", "topSites", "storage", "notifications", "contextMenus", "offscreen", "activeTab", "webview")
    "Low"      = @("alarms", "background", "idle", "fontSettings", "printing", "favicon", "unlimitedStorage", "system.cpu", "system.memory", "system.network", "ttsEngine", "telemetry", "mozillaAddons", "systemPrivate")
}

$hostRiskCategories = @{
    "Critical" = @("<all_urls>", "*://*/*", "file:///*", "*://*/", "https://*/*", "http://*/*", "*://*/*/*", "ftp://*/*")
    "High"     = @("*://github.com/*", "*://gmail.com/*", "*://outlook.com/*", "*://docs.google.com/*", "*://drive.google.com/*", "*://*.google.com/*", "*://microsoft.com/*", "*://*.microsoft.com/*", "*://office.com/*", "*://*.office.com/*", "*://sharepoint.com/*", "*://*.sharepoint.com/*", "*://onedrive.com/*", "*://*.onedrive.com/*", "*://azure.com/*", "*://*.azure.com/*", "*://aws.amazon.com/*", "*://*.aws.amazon.com/*", "*://console.cloud.google.com/*", "*://banking.com/*", "*://*.bank/*", "*://paypal.com/*", "*://stripe.com/*", "*://slack.com/*", "*://teams.microsoft.com/*", "*://zoom.us/*")
}

$builtInExtensions = @{

    # Edge
    "ahfgeienlihckogmohjhadlkjgocpleb" = "Web Store"
    "dgiklkfkllikcanfonkcabmbdfmgleag" = "Microsoft Clipboard Extension" 
    "epdpgaljfdjmcemiaplofbiholoaepem" = "Suppress Consent Prompt"
    "fikbjbembnmfhppjfnmfkahdhfohhjmg" = "Media Internals Services Extension"
    "jmjflgjpcpepeafmmgdpfkogkghcpiha" = "Edge relevant text changes"
    "kfbdpdaobnofkbopebjglnaadopfikhh" = "Microsoft Edge DevTools Enhancements"
    "ncbjelpjchkpbikbpkcchkhkblodoama" = "WebRTC Internals Extension"
    
    # Chrome
    "fignfifoniblkonapihmkfakmlgkbkcf" = "Google Network Speech"
    "ghbmnnjooekpmoecnnnilnnbdlolhkhi" = "Google Docs Offline"
    "mhjfbmdgcfjbbpaeojofohoefgiehjai" = "Chrome PDF Viewer"
    "nkeimhogjdpnpccoofpliimaahmaaome" = "Google Hangouts"
    "nmmhkkegccagdldgiimedpiccmgmieda" = "Chrome Web Store Payments"
    
    # Firefox
    "formautofill@mozilla.org"         = "Form Autofill"
    "newtab@mozilla.org"               = "New Tab"
    "pictureinpicture@mozilla.org"     = "Picture-In-Picture"
    "default-theme@mozilla.org"        = "System theme"
}

# Priority mapping
$priority = [ordered]@{
    "Critical" = 1
    "High"     = 2
    "Medium"   = 3
    "Low"      = 4
    "Unknown"  = 5
}

# Reverse lookup: API -> numeric risk
$apiToRiskMap = @{}
foreach ($risk in $priority.Keys) {
    foreach ($api in $apiRiskCategories[$risk]) {
        $apiToRiskMap[$api] = $priority[$risk]
    }
}

function Get-APIRiskLevel {
    param($apis)

    if (-not $apis -or $apis.Count -eq 0) { return "Unknown" }

    $minScore = 5
    foreach ($api in $apis) {
        if ($apiToRiskMap.ContainsKey($api)) {
            $score = $apiToRiskMap[$api]
            if ($score -lt $minScore) { $minScore = $score }
        }
    }

    return ($priority.Keys | Where-Object { $priority[$_] -eq $minScore })
}

function Get-HostRiskLevel {
    param($hostPermissions)
    if (-not $hostPermissions -or $hostPermissions.Count -eq 0) { return "Unknown" }

    $bestScore = 5
    foreach ($permission in $hostPermissions) {
        foreach ($risk in @("High","Critical")) {  # High first
            foreach ($pattern in $hostRiskCategories[$risk]) {
                if ($permission -eq $pattern) {
                    $score = $priority[$risk]
                    if ($score -lt $bestScore) {
                        $bestScore = $score
                    }
                }
            }
        }
    }
    return ($priority.Keys | Where-Object { $priority[$_] -eq $bestScore })
}


$results = @()

foreach ($entry in $browserPaths) {
    $pattern = Join-Path $BaseUserPath $entry.Pattern
    Write-Verbose "Scanning pattern: $pattern (Browser: $($entry.Browser))"

    Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | ForEach-Object {
        $profileFolder = Split-Path $_.FullName -Parent | Split-Path -Leaf

        if ([string]::IsNullOrWhiteSpace($entry.Regex) -or $profileFolder -match $entry.Regex) {
            Write-Verbose "  Processing profile folder: $profileFolder ($($entry.Browser))"

            $profileName = $profileFolder
            $profilePath = "C:\Users\{0}" -f ($_.FullName -split '\\')[2]
        
            try {
                $json = Get-Content $_.FullName -Raw | ConvertFrom-Json
                $extensions = @()
            
                # Normalize extensions based on browser type
                if ($json.addons) {

                    # Firefox format
                    $extensions = $json.addons | Where-Object { $_.active -and $_.type -eq "extension" } | ForEach-Object {
                        @{ ID               = $_.id
                            Name            = $_.defaultLocale.name
                            Version         = $_.version
                            APIs            = $_.userPermissions.permissions
                            HostPermissions = $_.userPermissions.origins
                        }
                    }
                }
                elseif ($json.extensions.settings) {

                    # Chromium format  
                    $extensions = $json.extensions.settings.PSObject.Properties | Where-Object { -not $_.Value.state -or $_.Value.state -eq 1 } | ForEach-Object {
                        if ($_.Value.manifest) {
                            @{ 
                                ID              = $_.Name
                                Name            = $_.Value.manifest.browser_action.default_title ?? $_.Value.manifest.name
                                Version         = $_.Value.manifest.version
                                APIs            = $_.Value.active_permissions.api
                                HostPermissions = @(
                                    $_.Value.active_permissions.explicit_host
                                    $_.Value.active_permissions.scriptable_host
                                    $_.Value.active_permissions.origins
                                ) | Where-Object { $_ }
                            }
                        }
                    } | Where-Object { $_ }
                }
            
                # Process extensions
                $extensions | ForEach-Object {
                    $isOutOfBox = $builtInExtensions.ContainsKey($_.ID)
                    $apiRisk = Get-APIRiskLevel $_.APIs
                    $hostRisk = Get-HostRiskLevel $_.HostPermissions
                
                    $results += [PSCustomObject]@{
                        Browser         = $entry.Browser
                        ProfilePath     = $profilePath  
                        Profile         = $profileName
                        ExtensionID     = $_.ID
                        Name            = $_.Name
                        Version         = $_.Version
                        OutOfBox        = $isOutOfBox
                        APIRisk         = $apiRisk
                        HostRisk        = $hostRisk
                        HostPermissions = ($_.HostPermissions -join ", ") ?? ""
                        APIPermissions  = ($_.APIs -join ", ") ?? ""
                    }
                }
            }
            catch {
                Write-Warning ("Failed to parse {0}: {1}" -f $_.FullName, $_.Exception.Message)
            }
        }
        else {
            Write-Verbose "  Skipping folder: $profileFolder (did not match regex $($entry.Regex))"
        }
    }
}

# Apply filters
Write-Verbose "Applying filters..."
Write-Verbose "  Browser: $Browser"
Write-Verbose "  RiskLevel: $RiskLevel"
Write-Verbose "  ExcludeBuiltIn: $ExcludeBuiltIn"
Write-Verbose "  IncludeBuiltInOnly: $IncludeBuiltInOnly"
Write-Verbose "  UserProfile: $UserProfile"
Write-Verbose "  Results before filtering: $($results.Count)"

if ($Browser) { $results = $results | Where-Object { $_.Browser -in $Browser } }
if ($RiskLevel) { $results = $results | Where-Object { $_.APIRisk -in $RiskLevel -or $_.HostRisk -in $RiskLevel } }
if ($ExcludeBuiltIn) { $results = $results | Where-Object { -not $_.OutOfBox } }
if ($IncludeBuiltInOnly) { $results = $results | Where-Object { $_.OutOfBox } }
if ($UserProfile) { $results = $results | Where-Object { $_.ProfilePath -in $UserProfile } }

Write-Verbose "  Results after filtering: $($results.Count)"

# Output
if ($ExportPath) {
    $results | Export-Csv -Path $ExportPath -NoTypeInformation
}
else {
    $sorted = $results |
    Sort-Object `
    @{ Expression = { $priority[$_.HostRisk] } }, `
    @{ Expression = { $priority[$_.APIRisk] } }, `
    @{ Expression = { $_.Browser } }, `
    @{ Expression = { $_.ProfilePath } }, `
    @{ Expression = { $_.Profile } }

    switch ($OutputFormat) {
        "Table" {
            $sorted |
            Select-Object Browser, Profile, Name, Version, APIRisk, HostRisk,
            @{ Name = 'HostPermissions'; Expression = { ($_.HostPermissions -join ', ') -replace '(.{80}).+', '$1...' } },
            @{ Name = 'APIPermissions'; Expression = { ($_.APIPermissions -join ', ') -replace '(.{80}).+', '$1...' } } |
            Format-Table -AutoSize
        }
        "List" {
            $sorted |
            Format-List Browser, ProfilePath, Profile, ExtensionID, Name, Version, OutOfBox,
            APIRisk, HostRisk, HostPermissions, APIPermissions
        }
    }
}