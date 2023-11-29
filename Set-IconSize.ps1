<#
.SYNOPSIS
This script is used to resize icons for applications in ConfigMgr. 
Large icons result in a larger base64 string which can cause issues with the size of the CI download by BITS from a CMG. 
This script will resize the icon to a smaller size in ConfigMgr.
Original and resized icons are saved in the $IconBackupDirectory

.NOTES
Author Ben Whitmore
Date: 2023-11-25
Version: 1.0

################# DISCLAIMER #################
The author provides scripts, macro, and other code examples for illustration only, without warranty 
either expressed or implied, including but not limited to the implied warranties of merchantability 
and/or fitness for a particular purpose. This script is provided 'AS IS' and the author does not 
guarantee that the following script, macro, or code can or should be used in any situation or that 
operation of the code will be error-free.

.PARAMETER SiteCode
The ConfigMgr site code

.PARAMETER ProviderMachineName
The ConfigMgr site server

.PARAMETER ApplicationName
The name of the application to resize the icon for. The default value is "*"

.PARAMETER NewWidth
The new width of the resized icon. The default value is 110.

.PARAMETER IconBackupDir
The directory path where the icon backup will be created and the resized icon will be saved. The default value is "C:\ConfigMgrIconBackup".

.EXAMPLE
Set-IconSize.ps1 -AppName "*Microsoft*" -NewWidth 110 -IconBackupDir "E:\IconBackup" -OnlyPMPApps
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SiteCode,
    [Parameter(Mandatory = $true, Position = 1)]
    [string]$ProviderMachineName,
    [Parameter(Mandatory = $false, Position = 2)]
    [string]$ApplicationName = "*",
    [Parameter(Mandatory = $false, Position = 3)]
    [int]$NewWidth = 110,
    [Parameter(Mandatory = $false, Position = 4)]
    [string]$IconBackupDir = "C:\ConfigMgrIconBackup",
    [Parameter(Mandatory = $false)]
    [switch]$OnlyPMPApps
)
function New-IconBackupDir {
    param (
        [Parameter(Mandatory = $false)]
        [string]$IconBackupDir
    )

    # Set icon backup directory name
    $iconDirectory = Join-Path -Path $IconBackupDir -ChildPath $dateToString
    try {

        # Create a new directory for icon backup
        If (-not (Test-Path -Path $iconDirectory) ) {
            Write-Verbose ("Creating icon backup directory '{0}'..." -f $iconDirectory) -Verbose
            New-Item -Path  $iconDirectory -ItemType Directory -Force | Out-Null
        } 
        else {
            Write-Verbose ("Directory '{0}' already exists" -f $iconDirectory) -Verbose
        }
    }
    catch {
        Write-Verbose ("Error creating directory '{0}'" -f $iconDirectory) -Verbose 
        throw
    }
}

function Connect-SiteServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteCode,
        [Parameter(Mandatory = $true)]
        [string]$ProviderMachineName
    )
    Write-Verbose -Message "Importing Module: ConfigurationManager.psd1 and connecting to Provider $($ProviderMachineName)..."

    # Import the ConfigurationManager.psd1 module 
    try {

        # Check if the ConfigurationManager module is already imported
        if (-not (Get-Module ConfigurationManager)) {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -Verbose:$False
        }
    }
    catch {
        Write-Verbose -Message 'Warning: Could not import the ConfigurationManager.psd1 Module'
        throw
    }

    # Check Provider is valid
    if (!($ProviderMachineName -eq (Get-PSDrive -ErrorAction SilentlyContinue | Where-Object { $_.Provider -like "*CMSite*" }).Root)) {
        Write-Verbose -Message ("Could not connect to the Provider '{0}'. Did you specify the correct Site Server?" -f $iconDirectory) -Verbose
        break
    }
    else {

        # Write which provider we are connecting to
        Write-Verbose -Message ("Connected to provider '{0}'" -f $ProviderMachineName) -Verbose
    }

    try {

        # Check if the site's drive is already present
        if (-not ( $SiteCode -eq (Get-PSDrive -ErrorAction SilentlyContinue | Where-Object { $_.Provider -like "*CMSite*" }).Name) ) {
            Write-Verbose -Message ("No PSDrive found for '{0}' in PSProvider CMSite for Root '{1}'. Did you specify the correct Site Code?" -f $SiteCode, $ProviderMachineName) -Verbose
            break
        }
        else {

            # Connect to the site's drive
            Write-Verbose -Message ("Connected to PSDrive " -f $SiteCode) -Verbose
            Set-Location "$($SiteCode):\"
        }
    }
    catch {
        Write-Verbose -Message ("Warning: Could not connect to the specified provider '{0}' at site '{1}'" -f $SiteCode, $ProviderMachineName) -Verbose
        break
    }
}

function Set-IconSizeLower {
    param (        
        [Parameter(Mandatory = $false)]
        [string[]]$ApplicationName,
        [Parameter(Mandatory = $false)]
        [string]$IconBackupDir = $IconBackupDir,
        [Parameter(Mandatory = $false)]
        [int]$NewWidth = $NewWidth
    )

    # Characters that are not allowed in the icon path name when saving icons
    $invalidChars = '[<>:"/\\|\?\*]'
        
    # Grab the applications
    if ($OnlyPMPApps) {
        $apps = Get-CMApplication -Name "*$ApplicationName*" | Where-Object {  $_.CIType_ID -eq 10 -and $_.IsLatest -eq $true -and $_.LocalizedDescription -like "Created by Patch My PC*" } 
    }
    else {
        $apps = Get-CMApplication -Name "*$ApplicationName*" | Where-Object { $_.CIType_ID -eq 10 -and $_.IsLatest -eq $true }
    }

    # Grab properties to display in OGV including current icon data length
    $appResults = $apps | ForEach-Object {
        $xml = [xml]$_.SDMPackageXML
        $singleIcon = $xml.AppMgmtDigest.Resources.Icon.Data
        $singleIconLength = $xml.AppMgmtDigest.Resources.Icon.Data.Length
        $singleIconStream = [System.IO.MemoryStream][System.Convert]::FromBase64String($singleIcon)
        $singleIconBitmap = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($singleIconStream)
        $_ | Select-Object -Property CI_ID, @{Name = 'AppName'; Expression = { $_.LocalizedDisplayName } }, @{Name = 'IconLength'; Expression = { $singleIconLength } }, @{Name = 'IconWidth'; Expression = { $singleIconBitmap.Width } }
    }

    $ogvResults = $appResults | Sort-Object IconLength -Descending | Out-GridView -Title "Select the Application(s) to resize the icon for" -PassThru
    
    # If a selection is made, resize the icon(s)
    if ($ogvResults) {

        # Create the icon backup directory if it doesn't exist
        New-IconBackupDir -IconBackupDir $IconBackupDir
    
        foreach ($app in $ogvResults) {

            # Loop through the selected applications
            Write-Host "`n"
            Write-Verbose -Message ("Getting details for Application '{0}' with CI_ID '{1}'" -f $app.AppName, $app.CI_ID) -Verbose
            $appXml = Get-CMApplication -ID $app.CI_ID | Where-Object { $Null -ne $_.SDMPackageXML } | Select-Object -ExpandProperty SDMPackageXML
            
            # Get the icon for the application
            $package = [xml]($appXml)
            $icon = $package.AppMgmtDigest.Resources.Icon.Data
            
            # Sanitize the folder names
            $ApplicationNameSanitized = ($app.AppName -replace $invalidChars, '_').TrimEnd('.', ' ')

            if ($icon) {

                # Convert the icon to a bitmap
                $iconStream = [System.IO.MemoryStream][System.Convert]::FromBase64String($icon)
                $iconBitmap = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($iconStream)
                Write-Verbose -Message ("Original icon size is {0}" -f $iconBitmap.Size) -Verbose
                Write-Verbose -Message ("Original icon base64 length is {0}" -f $app.IconLength) -Verbose

                if ($iconBitmap.Width -gt $NewWidth) {
                    Write-Verbose -Message ("Resizing icon for '{0}'" -f $app.AppName) -Verbose
                
                    # Backup the existing icon
                    $iconPath = Join-Path -Path $IconBackupDir -ChildPath "$($ApplicationNameSanitized).png"
                    $iconBitmap.Save($iconPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    Write-Verbose -Message ("Original icon saved to '{0}'" -f $iconPath) -Verbose

                    # Resize the icon
                    $newHeight = ($iconBitmap.Height / $iconBitmap.Width) * $newWidth
                    $newIconBitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $iconBitmap, $NewWidth, $newHeight
                    $newIconBitmap.SetResolution($iconBitmap.HorizontalResolution, $iconBitmap.VerticalResolution)
                    Write-Verbose -Message ("New icon size is {0}" -f $newIconBitmap.Size) -Verbose

                    # Save the new icon
                    $newIconPath = Join-Path -Path $IconBackupDir -ChildPath "$($ApplicationNameSanitized)_Resized.png"
                    $newIconBitmap.Save($newIconPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    Write-Verbose -Message ("New icon saved to '{0}'" -f $newIconPath) -Verbose

                    if (Test-Path -Path $newIconPath -PathType Leaf) {

                        # Read the bitmap file as a byte array
                        $imageBytes = [System.IO.File]::ReadAllBytes($newIconPath)
                
                        # Convert the byte array to base64
                        $base64String = [System.Convert]::ToBase64String($imageBytes)
                
                        # Get the length of the base64 string
                        $base64Length = $base64String.Length
                        Write-Verbose -Message ("New icon base64 length is '{0}'" -f $base64Length) -Verbose
                    }

                    # Update the application with the resized icon
                    try {
                        Set-CMApplication -Id $app.CI_ID -IconLocationFile $newIconPath
                        $newAppXml = Get-CMApplication -ID $app.CI_ID | Where-Object { $Null -ne $_.SDMPackageXML } | Select-Object -ExpandProperty SDMPackageXML
            
                        # Get the icon for the application
                        $newPackage = [xml]($newAppXml)
                        $newIcon = $newPackage.AppMgmtDigest.Resources.Icon.Data
                        if ($newIcon.Length -lt $app.IconLength) {
                            Write-Verbose -Message ("Icon for '{0}' was resized and updated in ConfigMgr successfully" -f $app.AppName) -Verbose
                        }
                        else {
                            Write-Verbose -Message ("Error updating icon for '{0}'. The icon was not updated in ConfigMgr." -f $app.AppName) -Verbose
                        }
                    }
                    catch {
                        Write-Verbose -Message ("Error updating icon for '{0}'" -f $app.AppName) -Verbose
                        throw
                    }
                }
                else {
                    Write-Verbose -Message ("Icon for '{0}' is already '{1}' the new width of '{2}'" -f $app.AppName, $(if ($iconBitmap.Width -eq $NewWidth) {'equal to'} else {'less than'}), $NewWidth) -Verbose
                    continue
                }
            }
            else {
                Write-Verbose -Message ("There is no custom icon assigned to '{0}'" -f $app.AppName) -Verbose
            }
        }
    } 
    else {
        Write-Verbose -Message "No applications selected" -Verbose
        return
    }
}

# Connect to the site server and set the icon size(s)
Connect-SiteServer -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName
if (-not (Test-Path -Path $IconBackupDir)) { New-IconBackupDir -IconBackupDir $IconBackupDir}
Set-IconSizeLower -ApplicationName $ApplicationName -IconBackupDir $IconBackupDir -NewWidth $NewWidth
Set-Location $PSScriptRoot