<#
.SYNOPSIS
    Enumerates browser extensions for Edge, Chrome, and Firefox for all user profiles on the system.
.DESCRIPTION
    Scans user profile directories for browser extension configuration files, parses extension metadata,
    evaluates risk based on permissions, and outputs a summary table.
.NOTES
    Author Ben Whitmore@PatchMyPC
    Created 22nd September 2025

    ################# DISCLAIMER #################
    Patch My PC provides scripts, macro, and other code examples for illustration only, without warranty 
    either expressed or implied, including but not limited to the implied warranties of merchantability 
    and/or fitness for a particular purpose. This script is provided 'AS IS' and Patch My PC does not 
    guarantee that the following script, macro, or code can or should be used in any situation or that 
    operation of the code will be error-free.
.EXAMPLE
    .\Get-BrowserExtensions.ps1
#>

$browserPaths = @{
    "Local\Microsoft\Edge\User Data\Default\Preferences"  = "Edge"
    "Local\Microsoft\Edge\User Data\Profile*\Preferences" = "Edge"
    "Local\Google\Chrome\User Data\Default\Preferences"   = "Chrome"
    "Local\Google\Chrome\User Data\Profile*\Preferences"  = "Chrome"
    "Roaming\Mozilla\Firefox\Profiles\*\extensions.json"  = "Firefox"
}

$VeryHigh = @("webRequest", "webRequestBlocking", "webRequestAuthProvider", "cookies", "nativeMessaging", "debugger", "scripting", "tabCapture", "management", "clipboardRead", "geolocation", "camera", "microphone", "desktopCapture")
$High = @("history", "downloads", "devtools". "clipboardWrite", "contentSettings", "proxy", "webNavigation", "privacy", "tabs", "identity", "pageCapture")
$Medium = @("bookmarks", "topSites", "storage", "notifications", "contextMenus", "offscreen", "activeTab", "webview")
$Low = @("alarms", "background", "idle", "fontSettings", "printing", "favicon", "unlimitedStorage", "system.cpu", "system.memory", "system.network", "ttsEngine", "telemetry", "mozillaAddons", "systemPrivate")

$BuiltInExtensions = @{

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

$results = @()

foreach ($pathPattern in $browserPaths.Keys) {
    $browserName = $browserPaths[$pathPattern]
    $fullPath = Join-Path -Path "C:\Users\*\AppData\" -ChildPath $pathPattern
    
    Get-ChildItem -Path $fullPath -ErrorAction SilentlyContinue | ForEach-Object {
        $file = $_
        $profileName = Split-Path (Split-Path $file.FullName) -Leaf
        $userName = ($file.FullName -split '\\')[2]
        $profilePath = "C:\Users\$userName"
        
        try {
            $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
            $extensions = @()
            if ($json.addons) {
                $extensions = $json.addons | Where-Object { $_.active -eq $true -and $_.type -eq "extension" } | ForEach-Object {
                    [PSCustomObject]@{
                        ID          = $_.id
                        Name        = $_.defaultLocale.name
                        Version     = $_.version
                        Permissions = $_.userPermissions.permissions
                    }
                }
            }
            elseif ($json.extensions.settings) {
                $extensions = $json.extensions.settings.PSObject.Properties | Where-Object { -not $_.Value.state -or $_.Value.state -eq 1 } | ForEach-Object {
                    if ($_.Value.manifest) {
                        [PSCustomObject]@{
                            ID          = $_.Name
                            Name        = if ($_.Value.manifest.browser_action.default_title) { $_.Value.manifest.browser_action.default_title } else { $_.Value.manifest.name }
                            Version     = $_.Value.manifest.version
                            Permissions = $_.Value.active_permissions.api
                        }
                    }
                } | Where-Object { $_ -ne $null }
            }
            $extensions | ForEach-Object {
                $risk = "N/A"
                
                if ($BuiltInExtensions.ContainsKey($_.ID)) {
                    $risk = "Out-of-box"
                }

                elseif ($_.Permissions | Where-Object { $VeryHigh -contains $_ }) { $risk = "Very High" }
                elseif ($_.Permissions | Where-Object { $High -contains $_ }) { $risk = "High" }
                elseif ($_.Permissions | Where-Object { $Medium -contains $_ }) { $risk = "Medium" }
                elseif ($_.Permissions | Where-Object { $Low -contains $_ }) { $risk = "Low" }
                
                $results += [PSCustomObject]@{
                    Browser     = $browserName
                    ProfilePath = $profilePath
                    Profile     = $profileName
                    ExtensionID = $_.ID
                    Name        = $_.Name
                    Version     = $_.Version
                    Risk        = $risk
                    APIs        = if ($_.Permissions) { $_.Permissions -join ", " } else { "" }
                }
            }
        }
        catch {
            Write-Warning "Failed to parse $($file.FullName): $($_.Exception.Message)"
        }
    }
}

$results | Format-Table -AutoSize