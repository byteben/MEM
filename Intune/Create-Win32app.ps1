<#
.Synopsis
Creates the Notepad++ Win32 app in Intune using the Graph SDK PowerShell module.

Created on:   16/08/2026
Created by:   Ben Whitmore@PatchMyPC
Filename:     Create-Win32app.ps1

.Description
- Creates the Notepad++ Win32 app in Intune using the Graph SDK PowerShell module. Downloads the Win32 Content Prep Tool and the Notepad++ installer, builds an .intunewin, then uploads and commits to Intune via Graph SDK.
- Requires Microsoft.Graph.Authentication, Az.Storage PowerShell modules. 
- Logs to console and CMTrace-compatible, log file in TEMP folder. Designed for testing and validation of Win32 app deployment in Intune.
- Client app used for authentication must have DeviceManagementApps.ReadWrite.All Graph API permission.

---------------------------------------------------------------------------------
LEGAL DISCLAIMER

The PowerShell script provided is shared with the community as-is
The author and co-author(s) make no warranties or guarantees regarding its functionality, reliability, or suitability for any specific purpose
Please note that the script may need to be modified or adapted to fit your specific environment or requirements
It is recommended to thoroughly test the script in a non-production environment before using it in a live or critical system
The author and co-author(s) cannot be held responsible for any damages, losses, or adverse effects that may arise from the use of this script
You assume all risks and responsibilities associated with its usage
---------------------------------------------------------------------------------

.Notes
- This script is intended for testing and validation purposes only and creates the Notepad++ 8.9.7 example used in the blog.
- Supply your Entra ID App Registration details as parameters when running the script.
- Application metadata, commands, detection, requirements, upload settings, and working paths are grouped in the CONFIGURATION region below.

.Example
.\Create-Win32app.ps1 -tenantId '<tenant-id>' -clientId '<client-id>' -clientSecret '<client-secret>'

.Example
.\Create-Win32app.ps1 -tenantId '<tenant-id>' -clientId '<client-id>' -certificateThumbprint '<certificate-thumbprint>'
#>

[CmdletBinding(DefaultParameterSetName = 'ClientSecret')]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$tenantId,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$clientId,

    [Parameter(Mandatory, ParameterSetName = 'ClientSecret')]
    [string]$clientSecret,

    [Parameter(Mandatory, ParameterSetName = 'Certificate')]
    [string]$certificateThumbprint
)

#region CONFIGURATION

# Application
$appName        = "Notepad++"
$appVersion     = "8.9.7"
$appPublisher   = "Notepad++ Team"
$appDescription = "Notepad++ is a free source code editor and Notepad replacement."
$informationUrl = "https://notepad-plus-plus.org/downloads/v8.9.7/"

# Source content
$installerUrl   = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.9.7/npp.8.9.7.Installer.x64.exe"
$contentPrepUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
$iconPath       = "$ENV:Temp\icon.png"

# Install / uninstall
$installArguments     = "/S"
$uninstallCommandLine = '"%ProgramFiles%\Notepad++\uninstall.exe" /S'

# Built-in requirements
$allowedArchitectures            = "x64"
$minimumSupportedOperatingSystem = @{ v10_1607 = $true }

# Detection
$registryDetectionRule = @{
    "@odata.type"        = "#microsoft.graph.win32LobAppRegistryRule"
    ruleType             = "detection"
    keyPath              = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++"
    valueName            = "DisplayVersion"
    operationType        = "version"
    operator             = "greaterThanOrEqual"
    comparisonValue      = "8.9.7.0"
    check32BitOn64System = $false
}

$rules = @(
    $registryDetectionRule
)

# Install experience / return codes
$installExperience = @{
    "@odata.type"         = "#microsoft.graph.win32LobAppInstallExperience"
    runAsAccount          = "system"
    deviceRestartBehavior = "suppress"
}

$returnCodes = @(
    @{ returnCode = 0;    type = "success" }
    @{ returnCode = 3010; type = "softReboot" }
    @{ returnCode = 1641; type = "hardReboot" }
    @{ returnCode = 1618; type = "retry" }
)

# Upload behaviour
$maxRetries                 = 10
$blockSize                  = 4 * 1024 * 1024
$sasRenewalThresholdMinutes = 5

# Working folders
$root          = Join-Path $env:TEMP "Win32App_NotepadPlusPlus"
$contentFolder = Join-Path $root "Content"
$outputFolder  = Join-Path $root "Output"
$toolFolder    = Join-Path $root "Tool"

#endregion

#region LOG
$logDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = Join-Path $env:TEMP ("Create-Win32app-{0}.log" -f $logDateTime)

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet(1, 2, 3)]
        [int]$Severity = 1,

        [Parameter(Mandatory = $false)]
        [string]$Component = "Create-Win32app"
    )

    $timeStamp = Get-Date -Format "HH:mm:ss.fff"
    $dateStamp = Get-Date -Format "MM-dd-yyyy"
    $logEntry = '<![LOG[{0}]LOG]!><time="{1}+000" date="{2}" component="{3}" context="" type="{4}" thread="" file="">' -f $Message, $timeStamp, $dateStamp, $Component, $Severity

    Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
}

function Write-LogAndHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet(1, 2, 3)]
        [int]$Severity = 1,

        [Parameter(Mandatory = $false)]
        [string]$Component = "Create-Win32app",

        [Parameter(Mandatory = $false)]
        [string]$ForegroundColor = "White"
    )

    switch ($Severity) {
        2 { $ForegroundColor = "Yellow" }
        3 { $ForegroundColor = "Red" }
    }

    Write-Host $Message -ForegroundColor $ForegroundColor
    Write-CMTraceLog -Message $Message -Severity $Severity -Component $Component
}

function Write-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Number,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $stageMessage = ("[{0}/8] {1}" -f $Number, $Message)
    Write-Host ""
    Write-Host $stageMessage -ForegroundColor Cyan
    Write-CMTraceLog -Message $stageMessage -Component "Create-Win32app"
}

function Get-GraphErrorDetail {
    <#
    .Description
    Extracts the full error detail from a failed Invoke-MgGraphRequest call.
    $_.Exception.Message only returns the HTTP status line. The actual Graph error
    code, message and request ID are in $_.ErrorDetails.Message as JSON.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $detail = [ordered]@{
        ExceptionMessage = $ErrorRecord.Exception.Message
        GraphErrorCode   = $null
        GraphMessage     = $null
        RequestId        = $null
        RawErrorBody     = $null
    }

    if ($ErrorRecord.ErrorDetails.Message) {
        $detail.RawErrorBody = $ErrorRecord.ErrorDetails.Message

        try {
            $parsed = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop

            if ($parsed.error) {
                $detail.GraphErrorCode = $parsed.error.code
                $detail.GraphMessage = $parsed.error.message
                $detail.RequestId = $parsed.error.innerError.'request-id'
            }
        }
        catch {
            # JSON parse failed - RawErrorBody already captured above
        }
    }

    return $detail
}

function Invoke-IntuneSasRenewal {
    <#
    .Description
    Requests a fresh Azure Storage SAS URI for an in-progress upload by calling the
    renewUpload action on the file entry, then polls until the renewal succeeds.
    The blob location (container/path) stays the same - only the SAS token is refreshed -
    so blocks already staged remain valid. Returns the new azureStorageUri.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileUri,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 10
    )

    Write-LogAndHost -Message "Renewing Azure Storage SAS URI..." -Severity 2
    Invoke-MgGraphRequest -Method POST -Uri ("{0}/renewUpload" -f $FileUri) -Body "{}" -ContentType "application/json" | Out-Null

    $attempt = 0
    do {
        Start-Sleep -Seconds 3
        $status = Invoke-MgGraphRequest -Method GET -Uri $FileUri
        $attempt++
        Write-LogAndHost -Message ("SAS renewal poll attempt {0}/{1}. State: '{2}'" -f $attempt, $MaxRetries, $status.uploadState) -ForegroundColor Cyan
    } until ($status.uploadState -eq 'azureStorageUriRenewalSuccess' -or $attempt -ge $MaxRetries)

    if ($status.uploadState -ne 'azureStorageUriRenewalSuccess') {
        throw ("SAS renewal failed after {0} attempts. Final state: '{1}'" -f $MaxRetries, $status.uploadState)
    }

    Write-LogAndHost -Message "SAS URI renewed successfully" -ForegroundColor Green
    return $status.azureStorageUri
}

function New-IntuneBlobClient {
    <#
    .Description
    Builds a CloudBlockBlob from the raw azureStorageUri STRING. The SAS token is passed
    to StorageCredentials as a string and the blob address is parsed without its query,
    so the SAS signature is never re-encoded/mangled by [System.Uri] - which is the usual
    cause of an immediate "Server failed to authenticate the request ... signature" 403.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$AzureStorageUri
    )

    $qIndex = $AzureStorageUri.IndexOf('?')
    if ($qIndex -lt 0) {
        throw "Azure Storage URI does not contain a SAS token."
    }

    $blobBaseUri = $AzureStorageUri.Substring(0, $qIndex)
    $sasToken    = $AzureStorageUri.Substring($qIndex + 1)   # SAS token without leading '?'

    $credentials = [Microsoft.Azure.Storage.Auth.StorageCredentials]::new($sasToken)
    return [Microsoft.Azure.Storage.Blob.CloudBlockBlob]::new([System.Uri]::new($blobBaseUri), $credentials)
}

function Get-StorageErrorDetail {
    <#
    .Description
    Extracts detail from an Azure Storage failure. The useful information (HTTP status,
    Azure error code, service request id, extended message) lives on the nested
    StorageException.RequestInformation, not on the top-level PowerShell wrapper message.
    Also builds the full inner-exception chain so nothing is lost.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $detail = [ordered]@{
        ExceptionChain    = $null
        HttpStatusCode    = $null
        HttpStatusMessage = $null
        StorageErrorCode  = $null
        ExtendedMessage   = $null
        ServiceRequestId  = $null
        AdditionalDetails = $null
        Summary           = $null
    }

    # Walk the inner-exception chain and locate any StorageException within it
    $chain = @()
    $ex = $ErrorRecord.Exception
    $storageEx = $null
    while ($ex) {
        $chain += ("[{0}] {1}" -f $ex.GetType().Name, $ex.Message)
        if (-not $storageEx -and $ex.GetType().FullName -like '*Storage.StorageException') {
            $storageEx = $ex
        }
        $ex = $ex.InnerException
    }
    $detail.ExceptionChain = $chain -join '  -->  '

    if ($storageEx -and $storageEx.RequestInformation) {
        $ri = $storageEx.RequestInformation
        $detail.HttpStatusCode    = $ri.HttpStatusCode
        $detail.HttpStatusMessage = $ri.HttpStatusMessage
        $detail.ServiceRequestId  = $ri.ServiceRequestID

        if ($ri.ExtendedErrorInformation) {
            $detail.StorageErrorCode = $ri.ExtendedErrorInformation.ErrorCode
            $detail.ExtendedMessage  = $ri.ExtendedErrorInformation.ErrorMessage
            if ($ri.ExtendedErrorInformation.AdditionalDetails -and $ri.ExtendedErrorInformation.AdditionalDetails.Count -gt 0) {
                $detail.AdditionalDetails = (($ri.ExtendedErrorInformation.AdditionalDetails.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join '; ')
            }
        }
    }

    $summaryParts = @()
    if ($detail.HttpStatusCode)   { $summaryParts += ("HTTP {0} {1}" -f [int]$detail.HttpStatusCode, $detail.HttpStatusMessage) }
    if ($detail.StorageErrorCode) { $summaryParts += ("Code={0}" -f $detail.StorageErrorCode) }
    if ($detail.ExtendedMessage)  { $summaryParts += $detail.ExtendedMessage }
    if ($detail.ServiceRequestId) { $summaryParts += ("x-ms-request-id={0}" -f $detail.ServiceRequestId) }
    if (-not $summaryParts)       { $summaryParts += $ErrorRecord.Exception.Message }
    $detail.Summary = $summaryParts -join ' | '

    return $detail
}

#endregion

Write-LogAndHost -Message ("Log started: {0}" -f $logPath) -ForegroundColor Cyan

#region MODULES
foreach ($module in @('Microsoft.Graph.Authentication', 'Az.Storage')) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-LogAndHost -Message ("Module '{0}' not found. Installing..." -f $module) -Severity 2
        Install-Module $module -Force -Scope CurrentUser
    }
    Import-Module $module -Force
    Write-LogAndHost -Message ("Module '{0}' loaded" -f $module) -ForegroundColor Cyan
}

#endregion

#region WORKING FOLDERS

if (-not (Test-Path $iconPath -PathType Leaf)) {
    throw ("Icon file not found: '{0}'" -f $iconPath)
}

# Preserve the icon bytes before recreating the temporary working folder.
$iconBytes = [IO.File]::ReadAllBytes($iconPath)

if (Test-Path $root) {
    Get-ChildItem $root -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

foreach ($folder in @($root, $contentFolder, $outputFolder, $toolFolder)) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

[IO.File]::WriteAllBytes((Join-Path $root "icon.png"), $iconBytes)

Write-LogAndHost -Message ("Working folder: '{0}'" -f $root) -ForegroundColor Cyan

#endregion

# Tracks which step we are in so the catch block can report exactly where it failed
$currentStep = "Initialise"

try {

    #region 1. DOWNLOAD WIN32 CONTENT PREP TOOL
    Write-Stage -Number 1 -Message "Build the .intunewin"
    $currentStep = "Download Win32 Content Prep Tool"
    $toolPath = Join-Path $toolFolder "IntuneWinAppUtil.exe"

    Write-LogAndHost -Message ("Downloading Win32 Content Prep Tool from '{0}'" -f $contentPrepUrl) -ForegroundColor Cyan

    Invoke-WebRequest -Uri $contentPrepUrl -OutFile $toolPath -UseBasicParsing

    if (-not (Test-Path $toolPath)) {
        throw ("Win32 Content Prep Tool download failed. File not found at '{0}'" -f $toolPath)
    }

    Write-LogAndHost -Message ("Win32 Content Prep Tool downloaded to '{0}'" -f $toolPath) -ForegroundColor Green

    #endregion

    #region 2. DOWNLOAD APPLICATION INSTALLER
    $currentStep = "Download application installer"

    $installerFileName = Split-Path $installerUrl -Leaf
    $installerPath = Join-Path $contentFolder $installerFileName

    Write-LogAndHost -Message ("Downloading '{0}' from '{1}'" -f $installerFileName, $installerUrl) -ForegroundColor Cyan

    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing

    if (-not (Test-Path $installerPath)) {
        throw ("Installer download failed. File not found at '{0}'" -f $installerPath)
    }

    Write-LogAndHost -Message ("Installer downloaded to '{0}'" -f $installerPath) -ForegroundColor Green

    #endregion

    #region 3. BUILD .intunewin
    $currentStep = "Build .intunewin"

    Write-LogAndHost -Message "Building .intunewin file..." -ForegroundColor Cyan

    $arguments = @('-s', ('"{0}"' -f $installerPath), '-c', ('"{0}"' -f $contentFolder), '-o', ('"{0}"' -f $outputFolder), '-q')

    Write-LogAndHost -Message ("Running: '{0} {1}'" -f $toolPath, ($arguments -join ' ')) -ForegroundColor Cyan

    $process = Start-Process -FilePath $toolPath -ArgumentList $arguments -Wait -PassThru

    if ($process.ExitCode -ne 0) {
        throw ("IntuneWinAppUtil.exe exited with code {0}" -f $process.ExitCode)
    }

    $intuneWinFile = Get-ChildItem $outputFolder -Filter "*.intunewin" | Select-Object -First 1

    if (-not $intuneWinFile) {
        throw ("No .intunewin file found in '{0}' after running the content prep tool" -f $outputFolder)
    }

    $intuneWinPath = $intuneWinFile.FullName
    Write-LogAndHost -Message ("Built .intunewin: '{0}'" -f $intuneWinPath) -ForegroundColor Green

    #endregion

    #region 4. EXTRACT .intunewin METADATA
    Write-Stage -Number 2 -Message "Read the .intunewin metadata and encryption information"
    $currentStep = "Extract .intunewin metadata"

    Write-LogAndHost -Message "Extracting .intunewin metadata..." -ForegroundColor Cyan

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $extractPath = Join-Path $root "Extract"

    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }

    New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($intuneWinPath, $extractPath)

    $metaXml = [xml](Get-Content (Join-Path $extractPath "IntuneWinPackage\Metadata\Detection.xml"))
    $encryptedFilePath = Get-ChildItem (Join-Path $extractPath "IntuneWinPackage\Contents") | Select-Object -First 1 -ExpandProperty FullName
    $sizeEncrypted = (Get-Item $encryptedFilePath).Length
    $sizeUnencrypted = [int64]$metaXml.ApplicationInfo.UnencryptedContentSize

    Write-LogAndHost -Message ("Metadata extracted. Encrypted: {0} bytes, Unencrypted: {1} bytes" -f $sizeEncrypted, $sizeUnencrypted) -ForegroundColor Green

    $encryptionInfo = @{
        encryptionKey        = $metaXml.ApplicationInfo.EncryptionInfo.EncryptionKey
        macKey               = $metaXml.ApplicationInfo.EncryptionInfo.MacKey
        initializationVector = $metaXml.ApplicationInfo.EncryptionInfo.InitializationVector
        mac                  = $metaXml.ApplicationInfo.EncryptionInfo.Mac
        profileIdentifier    = "ProfileVersion1"
        fileDigest           = $metaXml.ApplicationInfo.EncryptionInfo.FileDigest
        fileDigestAlgorithm  = $metaXml.ApplicationInfo.EncryptionInfo.FileDigestAlgorithm
    }

    Write-LogAndHost -Message ("Encryption info parsed. Algorithm: '{0}'" -f $encryptionInfo.fileDigestAlgorithm) -ForegroundColor Cyan

    #endregion

    #region 5. CONNECT TO GRAPH
    Write-Stage -Number 3 -Message "Connect to Microsoft Graph and create the Win32 app"
    $currentStep = "Connect to Microsoft Graph"

    # Disconnect any existing cached session
    Write-LogAndHost -Message "Disconnecting any existing Graph session..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
        Write-LogAndHost -Message "Connecting to Microsoft Graph using a certificate..." -ForegroundColor Cyan

        Connect-MgGraph `
            -TenantId $tenantId `
            -ClientId $clientId `
            -CertificateThumbprint $certificateThumbprint `
            -NoWelcome
    }
    else {
        Write-LogAndHost -Message "Connecting to Microsoft Graph using a client secret..." -ForegroundColor Cyan

        $secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

        Connect-MgGraph `
            -TenantId $tenantId `
            -ClientSecretCredential $credential `
            -NoWelcome
    }

    Write-LogAndHost -Message ("Connected to Microsoft Graph using {0} authentication" -f $PSCmdlet.ParameterSetName) -ForegroundColor Green

    #endregion

    #region 6. CREATE APP RECORD
    $currentStep = "Create app record"

    Write-LogAndHost -Message ("Creating app record for '{0}'..." -f $appName) -ForegroundColor Cyan

    $appBody = [ordered]@{
        "@odata.type"                   = "#microsoft.graph.win32LobApp"
        displayName                     = $appName
        description                     = $appDescription
        publisher                       = $appPublisher
        displayVersion                  = $appVersion
        informationUrl                  = $informationUrl

        largeIcon = @{
            "@odata.type" = "#microsoft.graph.mimeContent"
            type          = "image/png"
            value         = [Convert]::ToBase64String(
                [IO.File]::ReadAllBytes("$root\icon.png")
            )
        }

        fileName                        = $intuneWinFile.Name
        setupFilePath                   = $installerFileName
        installCommandLine              = "$installerFileName $installArguments"
        uninstallCommandLine            = $uninstallCommandLine
        allowedArchitectures            = $allowedArchitectures
        minimumSupportedOperatingSystem = $minimumSupportedOperatingSystem
        installExperience               = $installExperience
        returnCodes                     = $returnCodes
        rules                           = $rules
    } | ConvertTo-Json -Depth 10 -Compress

    $app = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -Body $appBody -ContentType "application/json"
    $appId = $app.id

    Write-LogAndHost -Message ("App record created. App ID: '{0}'" -f $appId) -ForegroundColor Green

    #endregion

    #region 7. CREATE CONTENT VERSION
    Write-Stage -Number 4 -Message "Create the content version and file entry"
    $currentStep = "Create content version"

    Write-LogAndHost -Message ("Creating content version for app '{0}'..." -f $appId) -ForegroundColor Cyan

    $contentVersion = Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/microsoft.graph.win32LobApp/contentVersions" -f $appId) -Body "{}" -ContentType "application/json"
    $contentVersionId = $contentVersion.id

    Write-LogAndHost -Message ("Content version created. Version ID: '{0}'" -f $contentVersionId) -ForegroundColor Green

    #endregion

    #region 8. CREATE FILE ENTRY
    $currentStep = "Create file entry"

    Write-LogAndHost -Message "Creating file entry..." -ForegroundColor Cyan

    $fileBody = [ordered]@{
        "@odata.type" = "#microsoft.graph.mobileAppContentFile"
        name          = $metaXml.ApplicationInfo.FileName
        size          = $sizeUnencrypted
        sizeEncrypted = $sizeEncrypted
        isDependency  = $false
        manifest      = $null
    } | ConvertTo-Json -Compress

    $fileEntry = Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/microsoft.graph.win32LobApp/contentVersions/{1}/files" -f $appId, $contentVersionId) -Body $fileBody -ContentType "application/json"
    $fileId = $fileEntry.id
    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/microsoft.graph.win32LobApp/contentVersions/{1}/files/{2}" -f $appId, $contentVersionId, $fileId

    Write-LogAndHost -Message ("File entry created. File ID: '{0}'" -f $fileId) -ForegroundColor Green

    #endregion

    #region 9. POLL FOR SAS URI
    Write-Stage -Number 5 -Message "Get and wait for the Azure Storage SAS URI"
    $currentStep = "Poll for SAS URI"

    Write-LogAndHost -Message "Polling for SAS URI..." -ForegroundColor Cyan

    $attempt = 0

    do {
        Start-Sleep -Seconds 3
        $fileStatus = Invoke-MgGraphRequest -Method GET -Uri $fileUri
        $attempt++
        Write-LogAndHost -Message ("SAS URI poll attempt {0}/{1}. State: '{2}'" -f $attempt, $maxRetries, $fileStatus.uploadState) -ForegroundColor Cyan
    } until ($fileStatus.uploadState -eq 'azureStorageUriRequestSuccess' -or $attempt -ge $maxRetries)

    if ($fileStatus.uploadState -ne 'azureStorageUriRequestSuccess') {
        throw ("Failed to get SAS URI after {0} attempts. Final state: '{1}'" -f $maxRetries, $fileStatus.uploadState)
    }

    $azureStorageUri = $fileStatus.azureStorageUri
    $sasUri = [System.Uri]::new($azureStorageUri)

    # Track SAS expiry so the upload loop can renew before it lapses
    $sasExpiry = [datetime]::MaxValue
    if ($fileStatus.azureStorageUriExpirationDateTime) {
        $sasExpiry = ([datetime]$fileStatus.azureStorageUriExpirationDateTime).ToUniversalTime()
        Write-LogAndHost -Message ("SAS URI obtained. Expires (UTC): {0:yyyy-MM-dd HH:mm:ss}" -f $sasExpiry) -ForegroundColor Green
    }
    else {
        Write-LogAndHost -Message "SAS URI obtained successfully" -ForegroundColor Green
    }

    #endregion

    #region 10. UPLOAD VIA CloudBlockBlob
    Write-Stage -Number 6 -Message "Upload the encrypted content to Azure Blob Storage"
    $currentStep = "Upload to Azure Blob Storage"

    $container = $sasUri.AbsolutePath.Split('/')[1]
    $blobPath = $sasUri.AbsolutePath.Substring($container.Length + 2)
    $totalBlocks = [Math]::Ceiling($sizeEncrypted / $blockSize)

    Write-LogAndHost -Message ("Starting upload. Container: '{0}', Blob: '{1}', Blocks: {2}" -f $container, $blobPath, $totalBlocks) -ForegroundColor Cyan

    $fileStream = [System.IO.File]::OpenRead($encryptedFilePath)
    $buffer = New-Object Byte[] $blockSize
    $blockIds = New-Object 'System.Collections.Generic.List[System.String]'
    $blobClient = New-IntuneBlobClient -AzureStorageUri $azureStorageUri

    try {
        $i = 0
        while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $encodedBlockId = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([Guid]::NewGuid().ToString()))
            $blockIds.Add($encodedBlockId)

            $memStream = New-Object System.IO.MemoryStream
            $memStream.Write($buffer, 0, $bytesRead)
            $memStream.Position = 0

            # Proactively renew the SAS if it is within 5 minutes of expiry
            if ((Get-Date).ToUniversalTime().AddMinutes($sasRenewalThresholdMinutes) -ge $sasExpiry) {
                Write-LogAndHost -Message "SAS URI is near expiry - renewing before next block" -Severity 2
                $azureStorageUri = Invoke-IntuneSasRenewal -FileUri $fileUri -MaxRetries $maxRetries
                $blobClient = New-IntuneBlobClient -AzureStorageUri $azureStorageUri
                $refreshed = Invoke-MgGraphRequest -Method GET -Uri $fileUri
                if ($refreshed.azureStorageUriExpirationDateTime) {
                    $sasExpiry = ([datetime]$refreshed.azureStorageUriExpirationDateTime).ToUniversalTime()
                }
            }

            # Upload the block. Retry with back-off; renew the SAS only if it is at/near
            # expiry - renewal cannot fix a corrupted signature, so we do not renew blindly.
            $blockAttempt = 0
            $blockUploaded = $false
            do {
                try {
                    $memStream.Position = 0
                    $blobClient.PutBlock($encodedBlockId, $memStream, $null)
                    $blockUploaded = $true
                }
                catch {
                    $blockAttempt++
                    $storageErr = Get-StorageErrorDetail -ErrorRecord $_
                    Write-LogAndHost -Message ("PutBlock failed for block {0} (attempt {1}/4). {2}" -f ($i + 1), $blockAttempt, $storageErr.Summary) -Severity 2

                    if ($blockAttempt -ge 4) {
                        $memStream.Dispose()
                        throw ("PutBlock failed for block {0} after {1} attempts. {2}" -f ($i + 1), $blockAttempt, $storageErr.Summary)
                    }

                    if ((Get-Date).ToUniversalTime().AddMinutes($sasRenewalThresholdMinutes) -ge $sasExpiry) {
                        try {
                            $azureStorageUri = Invoke-IntuneSasRenewal -FileUri $fileUri -MaxRetries $maxRetries
                            $blobClient = New-IntuneBlobClient -AzureStorageUri $azureStorageUri
                            $refreshed = Invoke-MgGraphRequest -Method GET -Uri $fileUri
                            if ($refreshed.azureStorageUriExpirationDateTime) {
                                $sasExpiry = ([datetime]$refreshed.azureStorageUriExpirationDateTime).ToUniversalTime()
                            }
                        }
                        catch {
                            Write-LogAndHost -Message ("SAS renewal attempt failed: {0}" -f $_.Exception.Message) -Severity 2
                        }
                    }

                    $backoff = [int][Math]::Min(30, [Math]::Pow(2, $blockAttempt))
                    Write-LogAndHost -Message ("Retrying block {0} in {1}s..." -f ($i + 1), $backoff) -Severity 2
                    Start-Sleep -Seconds $backoff
                }
            } until ($blockUploaded)

            $memStream.Dispose()

            $i++
            Write-LogAndHost -Message ("Uploaded block {0} of {1}" -f $i, $totalBlocks) -ForegroundColor Cyan
        }

        $blobClient.PutBlockList($blockIds)
        Write-LogAndHost -Message "All blocks committed to Azure Storage successfully" -ForegroundColor Green
    }
    finally {
        $fileStream.Dispose()
    }

    #endregion

    #region 11. VERIFY UPLOAD STATE BEFORE COMMIT
    $currentStep = "Verify upload state"

    Write-LogAndHost -Message "Verifying upload state before committing..." -ForegroundColor Cyan

    $attempt = 0

    do {
        Start-Sleep -Seconds 3
        $fileStatus = Invoke-MgGraphRequest -Method GET -Uri $fileUri
        $attempt++
        Write-LogAndHost -Message ("Upload state verification attempt {0}/{1}. State: '{2}'" -f $attempt, $maxRetries, $fileStatus.uploadState) -ForegroundColor Cyan
    } until ($fileStatus.uploadState -eq 'azureStorageUriRequestSuccess' -or $attempt -ge $maxRetries)

    if ($fileStatus.uploadState -ne 'azureStorageUriRequestSuccess') {
        throw ("Upload state check failed before commit after {0} attempts. Final state: '{1}'" -f $maxRetries, $fileStatus.uploadState)
    }

    Write-LogAndHost -Message ("Upload state confirmed '{0}'. Safe to commit" -f $fileStatus.uploadState) -ForegroundColor Green

    #endregion

    #region 12. COMMIT THE FILE
    Write-Stage -Number 7 -Message "Commit the uploaded content to Intune"
    $currentStep = "Commit file"

    Write-LogAndHost -Message "Committing file content to Intune..." -ForegroundColor Cyan

    $commitBody = @{ fileEncryptionInfo = $encryptionInfo } | ConvertTo-Json -Depth 5 -Compress

    Invoke-MgGraphRequest -Method POST -Uri ("{0}/commit" -f $fileUri) -Body $commitBody -ContentType "application/json"
    Write-LogAndHost -Message "Commit request sent successfully" -ForegroundColor Green

    #endregion

    #region 13. POLL FOR commitFileSuccess
    $currentStep = "Poll for commit completion"

    Write-LogAndHost -Message "Polling for commit completion..." -ForegroundColor Cyan

    $attempt = 0
    $success = $false

    do {
        Start-Sleep -Seconds 5
        $fileStatus = Invoke-MgGraphRequest -Method GET -Uri $fileUri
        $attempt++

        if ($fileStatus.uploadState -eq 'commitFileSuccess') {
            $success = $true
            Write-LogAndHost -Message ("Commit successful. State: '{0}'" -f $fileStatus.uploadState) -ForegroundColor Green
        }
        elseif ($fileStatus.uploadState -eq 'commitFileFailed') {
            throw ("Commit failed after {0} attempts. State: '{1}'" -f $attempt, $fileStatus.uploadState)
        }
        else {
            Write-LogAndHost -Message ("Commit poll attempt {0}/{1}. State: '{2}'" -f $attempt, $maxRetries, $fileStatus.uploadState) -Severity 2
        }
    } until ($success -or $attempt -ge $maxRetries)

    if (-not $success) {
        throw ("Commit did not complete within {0} attempts. Final state: '{1}'" -f $maxRetries, $fileStatus.uploadState)
    }

    #endregion

    #region 14. PATCH APP WITH COMMITTED CONTENT VERSION
    Write-Stage -Number 8 -Message "Point the app at the committed content version"
    $currentStep = "Patch app with committed content version"

    Write-LogAndHost -Message ("Updating app '{0}' with committed content version '{1}'..." -f $appId, $contentVersionId) -ForegroundColor Cyan

    $patchBody = @{
        "@odata.type"           = "#microsoft.graph.win32LobApp"
        committedContentVersion = $contentVersionId
    } | ConvertTo-Json -Compress

    Invoke-MgGraphRequest -Method PATCH -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}" -f $appId) -Body $patchBody -ContentType "application/json"

    $intuneAdminCenterUrl = "https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/2/appId/$appId"

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Win32 app creation complete" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ("App Name : {0}" -f $appName)
    Write-Host ("Version  : {0}" -f $appVersion)
    Write-Host ("App ID   : {0}" -f $appId)
    Write-Host ("Log file : {0}" -f $logPath)
    Write-Host ""
    Write-Host "Intune Admin Center:" -ForegroundColor Cyan
    Write-Host $intuneAdminCenterUrl -ForegroundColor Cyan

    Write-CMTraceLog -Message "Win32 app creation complete" -Component "Create-Win32app"
    Write-CMTraceLog -Message ("App Name: {0}; Version: {1}; App ID: {2}" -f $appName, $appVersion, $appId) -Component "Create-Win32app"
    Write-CMTraceLog -Message ("Intune Admin Center: {0}" -f $intuneAdminCenterUrl) -Component "Create-Win32app"

    #endregion
}
catch {

    Write-LogAndHost -Message ("Script failed at step: '{0}'" -f $currentStep) -Severity 3
    $graphError = Get-GraphErrorDetail -ErrorRecord $_
    $storageError = Get-StorageErrorDetail -ErrorRecord $_

    Write-LogAndHost -Message ("Exception     : {0}" -f $graphError.ExceptionMessage) -Severity 3

    if ($storageError.ExceptionChain) {
        Write-LogAndHost -Message ("Full detail   : {0}" -f $storageError.ExceptionChain) -Severity 3
    }

    if ($graphError.GraphErrorCode) {
        Write-LogAndHost -Message ("Graph code    : {0}" -f $graphError.GraphErrorCode) -Severity 3
    }

    if ($graphError.GraphMessage) {
        Write-LogAndHost -Message ("Graph message : {0}" -f $graphError.GraphMessage) -Severity 3
    }

    if ($graphError.RequestId) {
        Write-LogAndHost -Message ("Request ID    : {0}" -f $graphError.RequestId) -Severity 3
    }

    if ($graphError.RawErrorBody -and -not $graphError.GraphErrorCode) {
        Write-LogAndHost -Message ("Raw error body: {0}" -f $graphError.RawErrorBody) -Severity 3
    }

    if ($storageError.HttpStatusCode) {
        Write-LogAndHost -Message ("Storage HTTP  : {0} {1}" -f [int]$storageError.HttpStatusCode, $storageError.HttpStatusMessage) -Severity 3
    }

    if ($storageError.StorageErrorCode) {
        Write-LogAndHost -Message ("Storage code  : {0}" -f $storageError.StorageErrorCode) -Severity 3
    }

    if ($storageError.ExtendedMessage) {
        Write-LogAndHost -Message ("Storage msg   : {0}" -f $storageError.ExtendedMessage) -Severity 3
    }

    if ($storageError.ServiceRequestId) {
        Write-LogAndHost -Message ("x-ms-request-id : {0}" -f $storageError.ServiceRequestId) -Severity 3
    }

    if ($storageError.AdditionalDetails) {
        Write-LogAndHost -Message ("Storage detail: {0}" -f $storageError.AdditionalDetails) -Severity 3
    }

    Write-LogAndHost -Message ("Log file      : {0}" -f $logPath) -Severity 3

    throw
}
finally {
    if (Test-Path $root) {
        #Remove-Item $root -Recurse -Force
        Write-LogAndHost -Message ("Working folder '{0}' removed" -f $root) -ForegroundColor Cyan
    }
}