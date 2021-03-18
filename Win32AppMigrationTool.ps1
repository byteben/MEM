<#
.Synopsis
Created on:   14/03/2021
Created by:   Ben Whitmore
Filename:     Win32AppMigrationTool.ps1

The Win32 App Migration Tool is designed to inventory ConfigMgr Applications and Deployment Types, build .intunewin files and create Win3Apps in The MEM Admin Center

.Description
Version 1.0 - 14/03/2021

.Parameter AppName
Pass a string to the toll to search for applications in ConfigMgr

.Parameter SiteCode
Specify the Sitecode you wish to connect to

.Parameter ProviderMachineName
Specify the Site Server to connect to

.Parameter ExportLogo
When passed, the Application logo is decoded from base64 and saved to the Logos folder

.Parameter WorkingFolder
This is the working folder for the Win32AppMigration Tool. Care should be given when specifying the working folder because downloaded content can increase the working folder size considerably. The Following folders are created in this directory:-

-Content
-ContentPrepTool
-Details
-Logos
-Logs
-Win32Apps

.Parameter PackageApp
Pass this parameter to package selected apps in the .intunewin format

.Parameter CreateApps
Pass this parameter to create the Win32apps in Intune

.Example
.\Win32AppMigrationTool.ps1 -SiteCode "BB1" -ProviderMachineName "SCCM1.byteben.com" -AppName "Microsoft Edge Chromium *"

.Example
.\Win32AppMigrationTool.ps1 -SiteCode "BB1" -ProviderMachineName "SCCM1.byteben.com" -AppName "Microsoft Edge Chromium *" -ExportLogo

.Example
.\Win32AppMigrationTool.ps1 -SiteCode "BB1" -ProviderMachineName "SCCM1.byteben.com" -AppName "Microsoft Edge Chromium *" -ExportLogo -PackageApps

.Example
.\Win32AppMigrationTool.ps1 -SiteCode "BB1" -ProviderMachineName "SCCM1.byteben.com" -AppName "Microsoft Edge Chromium *" -ExportLogo -PackageApps -CreateApps

#>

Param (
    [Parameter(Mandatory = $True)]
    [String]$AppName,
    [String]$SiteCode,
    [String]$ProviderMachineName,
    [Parameter(Mandatory = $False)]
    [Switch]$ExportLogo,
    [String]$WorkingFolder = "C:\Win32AppMigrationTool",
    [Switch]$PackageApps,
    [Switch]$CreateApps
)

#Create Global Variables
$Global:WorkingFolder_Root = $WorkingFolder
$Global:WorkingFolder_Logos = Join-Path -Path $WorkingFolder_Root -ChildPath "Logos"
$Global:WorkingFolder_Content = Join-Path -Path $WorkingFolder_Root -ChildPath "Content"
$Global:WorkingFolder_ContentPrepTool = Join-Path -Path $WorkingFolder_Root -ChildPath "ContentPrepTool"
$Global:WorkingFolder_Logs = Join-Path -Path $WorkingFolder_Root -ChildPath "Logs"
$Global:WorkingFolder_Detail = Join-Path -Path $WorkingFolder_Root -ChildPath "Details"
$Global:WorkingFolder_Win32Apps = Join-Path -Path $WorkingFolder_Root -ChildPath "Win32Apps"

Function New-IntuneWin {
    Param (
        [String]$ContentFolder,
        [String]$OutputFolder,
        [String]$SetupFile
    )
    <#
    Function to create new Intunewin
    #>

    #If PowerShell is reference, grab the name of the .ps1 referenced in the Install Command line
    If ($SetupFile -match "powershell" -and $SetupFile -match "\.ps1") {
        Write-Host "Powershell script detected" -ForegroundColor Yellow
        Write-Host ''
        $Right = ($SetupFile -split ".ps1")[0]
        $Right = ($Right -Split " ")[-1]
        $Right = $Right.TrimStart("\", ".", "`"")
        $Command = $Right + ".ps1"
        Write-Host "Extracting the SetupFile Name for the Microsoft Win32 Content Prep Tool from the Install Command..." -ForegroundColor Cyan
        Write-Host $Command -ForegroundColor Green
    }
    else {

        #Search the Install Command line for other .exe installers
        If ($SetupFile -match "`.exe") {
            $Installer = ".exe"
            Write-Host "$Installer installer detected"
            $Right = ($SetupFile -split "\.exe")[0]
            $Right = ($Right -Split " ")[-1]
            $Command = $Right + $Installer
            Write-Host "Extracting the SetupFile Name for the Microsoft Win32 Content Prep Tool from the Install Command..." -ForegroundColor Cyan
            Write-Host $Command -ForegroundColor Green
        }
        elseif ($SetupFile -match "`.msi") {
            $Installer = ".msi"
            Write-Host "$Installer installer detected"
            $Right = ($SetupFile -split "\.msi")[0]
            $Right = ($Right -Split " ")[-1]
            $Command = $Right + $Installer
            Write-Host "Extracting the SetupFile Name for the Microsoft Win32 Content Prep Tool from the Install Command..." -ForegroundColor Cyan
            Write-Host $Command -ForegroundColor Green
        }
        elseif ($SetupFile -match "`.cmd") {
            $Installer = ".cmd"
            Write-Host "$Installer installer detected"
            $Right = ($SetupFile -split "\.cmd")[0]
            $Right = ($Right -Split " ")[-1]
            $Command = $Right + $Installer
            Write-Host "Extracting the SetupFile Name for the Microsoft Win32 Content Prep Tool from the Install Command..." -ForegroundColor Cyan
            Write-Host $Command -ForegroundColor Green
        }
        elseif ($SetupFile -match "`.bat") {
            $Installer = ".bat"
            Write-Host "$Installer installer detected"
            $Right = ($SetupFile -split "\.bat")[0]
            $Right = ($Right -Split " ")[-1]
            $Command = $Right + $Installer
            Write-Host "Extracting the SetupFile Name for the Microsoft Win32 Content Prep Tool from the Install Command..." -ForegroundColor Cyan
            Write-Host $Command -ForegroundColor Green
        }
    }
    Write-Host ''

    Try {
        #Check IntuneWinAppUtil.exe
        Write-Host "Re-checking presence of Win32 Content Prep Tool..." -ForegroundColor Cyan
        If (Test-Path (Join-Path -Path $WorkingFolder_ContentPrepTool -ChildPath "IntuneWinAppUtil.exe")) {
            Write-Host "Information: IntuneWinAppUtil.exe already exists at ""$($WorkingFolder_ContentPrepTool)"". Skipping download" -ForegroundColor Magenta
        }
        else {
            Write-Host "Downloading Win32 Content Prep Tool..." -ForegroundColor Cyan
            Get-FileFromInternet -URI "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" -Destination $WorkingFolder_ContentPrepTool
        }
        Write-Host ''
        Write-Host "Building IntuneWinAppUtil.exe execution string..." -ForegroundColor Cyan
        Write-Host """$($WorkingFolder_ContentPrepTool)\IntuneWinAppUtil.exe"" -s ""$($Command)"" -c ""$($ContentFolder)"" -o ""$($OutputFolder)""" -ForegroundColor Green

        #Try running the content prep tool to build the intunewin
        Try {
            $Arguments = @(
                "-s"
                $Command
                "-c"
                $ContentFolder
                "-o"
                $OutputFolder
                "-q"
            )
            Start-Process -FilePath (Join-Path -Path $WorkingFolder_ContentPrepTool -ChildPath "IntuneWinAppUtil.exe") -ArgumentList $Arguments -Wait
            Return $Right
        }
        Catch {
            Write-Host "Error creating the .intunewin file" -ForegroundColor Red
            Write-Host $_ -ForegroundColor Red
        }
    }
    Catch {
        Write-Host "The script encounted an error getting the Win32 Content Prep Tool" -ForegroundColor Red
    }
}
Function Get-ContentFiles {
    Param (
        [String]$Source,
        [String]$Destination
    )
    <#
    Function to download Deployment Type Content from Content Source Folder
    #>

    Try {
        $Robo = Robocopy.exe $Source $Destination /mir /e /z /r:5 /w:1 /reg /v /NDL /NJH /NJS /nc /ns /np
        $Robo 
        Return $Done
    }

    Catch {
        Write-Host "Error: Could not transfer content from ""$($Source)"" to ""$($Destination)"""
    }

}

Function Connect-SiteServer {
    Param (
        [String]$SiteCode,
        [String]$ProviderMachineName
    )
    <#
    Function to connect to ConfigMgr
    #>

    # Import the ConfigurationManager.psd1 module 
    Try {
        If ($Null -eq (Get-Module ConfigurationManager)) {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
        }
    }
    Catch {
        Write-Host 'Warning: Could not import the ConfigurationManager.psd1 Module' -ForegroundColor Red
    }

    # Connect to the site's drive if it is not already present
    Try {
        if ($Null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
        }
        #Set the current location to be the site code.
        Set-Location "$($SiteCode):\"
        Write-Host "Connected to provider ""$($ProviderMachineName)"" at site ""$($SiteCode)""" -ForegroundColor Green
    }
    Catch {
        Write-Host "Warning: Could not connect to the specified provider ""$($ProviderMachineName)"" at site ""$($SiteCode)""" -ForegroundColor Red
    }
    
}

Function New-FolderToCreate {
    <#
    Function to create folder structure for Win32AppMigrationTool
    #>  
    Param(
        [String]$Root,
        [String[]]$Folders
    )
    If (!($Root)) {
        Write-Host "Error: No Root Folder passed to Function" -ForegroundColor Red
    }
    If (!($Folders)) {
        Write-Host "Error: No Folder(s) passed to Function" -ForegroundColor Red
    }

    ForEach ($Folder in $Folders) {
        #Create Folders
        $FolderToCreate = Join-Path -Path $Root -ChildPath $Folder
        If (!(Test-Path $FolderToCreate)) {
            Write-Host "Creating Folder ""$($FolderToCreate)""..." -ForegroundColor Cyan
            Try {
                New-Item -Path $FolderToCreate -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "Folder ""$($FolderToCreate)"" created succesfully"
            }
            Catch {
                Write-Host "Warning: Couldn't create ""$($FolderToCreate)"" folder" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Information: Folder ""$($FolderToCreate)"" already exsts. Skipping folder creation" -ForegroundColor Magenta
        }
    }
}  

Function Export-Logo {
    <#
    Function to decode and export Base64 image for application logo to an output folder
    #>
    Param (
        [String]$IconId,
        [String]$AppName
    )
    Write-Host "Preparing to export Application Logo for ""$($AppName)"""
    If ($IconId) {

        #Check destination folder exists for logo
        If (!(Test-Path $WorkingFolder_Logos)) {
            Try {
                New-Item -Path $WorkingFolder_Logos -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            Catch {
                Write-Host "Warning: Couldn't create ""$($WorkingFolder_Logos)"" folder for Application Logos" -ForegroundColor Red
            }
        }

        #Continue if Logofolder exists
        If (Test-Path $WorkingFolder_Logos) {
            $LogoFolder_Id = (Join-Path -Path $WorkingFolder_Logos -ChildPath $IconId)
            $Logo_File = (Join-Path -Path $LogoFolder_Id -ChildPath Logo.jpg)

            #Continue if logo does not already exist in destination folder
            If (!(Test-Path $Logo_File)) {

                If (!(Test-Path $LogoFolder_Id)) {
                    Try {
                        New-Item -Path $LogoFolder_Id -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    }
                    Catch {
                        Write-Host "Warning: Couldn't create ""$($LogoFolder_Id)"" folder for Application Logo " -ForegroundColor Red
                    }
                }

                #Continue if Logofolder\<IconId> exists
                If (Test-Path $LogoFolder_Id) {
                    Try {
                        #Grab the SDMPackgeXML which contains the Application and Deployment Type details
                        $XMLPackage = Get-CMApplication -Name $AppName | Where-Object { $Null -ne $_.SDMPackageXML } | Select-Object -ExpandProperty SDMPackageXML

                        #Deserialize SDMPackageXML
                        $XMLContent = [xml]($XMLPackage)

                        $Raw = $XMLContent.AppMgmtDigest.Resources.icon.Data
                        $Logo = [Convert]::FromBase64String($Raw)
                        [System.IO.File]::WriteAllBytes($Logo_File, $Logo)
                        If (Test-Path $Logo_File) {
                            Write-Host "Success: Application logo for ""$($AppName)"" exported successfully to ""$($Logo_File)""" -ForegroundColor Green
                        }
                    }
                    Catch {
                        Write-Host "Warning: Could not export Logo to folder ""$($LogoFolder_Id)""" -ForegroundColor Red
                    }
                }
            }
            else {
                Write-Host "Information: Did not export Logo for ""$($AppName)"" to ""$($Logo_File)"" because the file already exists" -ForegroundColor Magenta
            }
        }
    }
    else {
        Write-Host "Warning: Null or invalid IconId passed to function. Could not export Logo" -ForegroundColor Red
    }
}

Function Get-FileFromInternet {
    <#
    Function to download and extract ContentPrep Tool
    #>
    Param (
        [String]$URI,
        [String]$Destination
    )
    
    $File = $URI -replace '.*/'
    $FileDestination = Join-Path -Path $Destination -ChildPath $File
    Try {
        Invoke-WebRequest -UseBasicParsing -Uri $URI -OutFile $FileDestination -ErrorAction Stop
    }
    Catch {
        Write-Host "Warning: Error downloading the Win32 Content Prep Tool" -ForegroundColor Red
        $_
    }
}
Function Get-AppInfo {
    <#
    Function to get deployment type(s) for applcation(s) passed
    #>
    Param (
        [String[]]$ApplicationName
    )

    #Create Array to display Application and Deployment Type Information
    $DeploymentTypes = @()
    $ApplicationTypes = @()
    $Content = @()

    #Iterate through each Application and get details
    ForEach ($Application in $ApplicationName) {

        #Grab the SDMPackgeXML which contains the Application and Deployment Type details
        $XMLPackage = Get-CMApplication -Name $Application | Where-Object { $Null -ne $_.SDMPackageXML } | Select-Object -ExpandProperty SDMPackageXML

        #Deserialize SDMPackageXML
        $XMLContent = [xml]($XMLPackage)

        #Get total number of Deployment Types for the Application
        $TotalDeploymentTypes = $XMLContent.AppMgmtDigest.Application.DeploymentTypes.DeploymentType.Count

        If ($TotalDeploymentTypes -gt 1) {

            $ApplicationObject = New-Object PSCustomObject
                
            #Application Details
            $ApplicationObject | Add-Member NoteProperty -Name Application_LogicalName -Value $XMLContent.AppMgmtDigest.Application.LogicalName
            $ApplicationObject | Add-Member NoteProperty -Name Application_Name -Value $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Title
            $ApplicationObject | Add-Member NoteProperty -Name Application_Description -Value $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Description
            $ApplicationObject | Add-Member NoteProperty -Name Application_Publisher -Value $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Publisher
            $ApplicationObject | Add-Member NoteProperty -Name Application_Version -Value $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Version
            $ApplicationObject | Add-Member NoteProperty -Name Application_IconId -Value $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Icon.Id
            $ApplicationObject | Add-Member NoteProperty -Name Application_TotalDeploymentTypes -Value $TotalDeploymentTypes
                
            #If we have the logo, add the path
            If (Test-Path -Path (Join-Path -Path $WorkingFolder_Logos -ChildPath (Join-Path -Path $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Icon.Id -ChildPath "Logo.jpg"))) {
                $ApplicationObject | Add-Member NoteProperty -Name Application_IconPath -Value (Join-Path -Path $WorkingFolder_Logos -ChildPath (Join-Path -Path $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Icon.Id -ChildPath "Logo.jpg"))
            }
            else {
                $ApplicationObject | Add-Member NoteProperty -Name Application_IconPath -Value $Null
            }
            $ApplicationTypes += $ApplicationObject
        
            #If Deployment Types exist, iterate through each DeploymentType and build deployment detail
            ForEach ($Object in $XMLContent.AppMgmtDigest.DeploymentType) {

                #Create new custom PSObjects to build line detail
                $DeploymentObject = New-Object -TypeName PSCustomObject
                $ContentObject = New-Object -TypeName PSCustomObject
                
                #DeploymentType Details
                $DeploymentObject | Add-Member NoteProperty -Name Application_LogicalName -Value $XMLContent.AppMgmtDigest.Application.LogicalName
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_LogicalName -Value $Object.LogicalName
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_Name -Value $Object.Title.InnerText
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_Technology -Value $Object.Installer.Technology
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_ExecutionContext -Value $Object.Installer.ExecutionContext
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_InstallContent -Value $Object.Installer.CustomData.InstallContent.ContentId
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_InstallCommandLine -Value $Object.Installer.CustomData.InstallCommandLine
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_UnInstallSetting -Value $Object.Installer.CustomData.UnInstallSetting
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_UninstallContent -Value $Object.Installer.CustomData.UninstallContent.ContentId
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_UninstallCommandLine -Value $Object.Installer.CustomData.UninstallCommandLine
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_ExecuteTime -Value $Object.Installer.CustomData.ExecuteTime
                $DeploymentObject | Add-Member NoteProperty -Name DeploymentType_MaxExecuteTime -Value $Object.Installer.CustomData.MaxExecuteTime

                $DeploymentTypes += $DeploymentObject

                #Content Details
                $ContentObject | Add-Member NoteProperty -Name Content_DeploymentType_LogicalName -Value $Object.LogicalName
                $ContentObject | Add-Member NoteProperty -Name Content_Location -Value $Object.Installer.Contents.Content.Location

                $Content += $ContentObject                
            }
        }
    } 
    Return $DeploymentTypes, $ApplicationTypes, $Content
}

Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Script Start Win32AppMigrationTool' -ForegroundColor DarkGray
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

#Connect to Site Server
Write-Host 'Connecting to Site Server..' -ForegroundColor Cyan
Connect-SiteServer -SiteCode  $SiteCode -ProviderMachineName $ProviderMachineName

#Region Check_Folders
Write-Host ''
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Checking Win32AppMigrationTool Folder Structure...' -ForegroundColor DarkGray
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

#Create Folders
Write-Host "Creating Folders..."-ForegroundColor Cyan
New-FolderToCreate -Root $WorkingFolder_Root -Folders @("", "Logos", "Content", "ContentPrepTool", "Logs", "Details", "Win32Apps")
#EndRegion Check_Folders

#Region Get_Content_Tool
Write-Host ''
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Checking Win32AppMigrationTool Content Tool...' -ForegroundColor DarkGray
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

#Download Win32 Content Prep Tool
If ($PackageApps) {
    Write-Host "Downloadling Win32 Content Prep Tool..." -ForegroundColor Cyan
    If (Test-Path (Join-Path -Path $WorkingFolder_ContentPrepTool -ChildPath "IntuneWinAppUtil.exe")) {
        Write-Host "Information: IntuneWinAppUtil.exe already exists at ""$($WorkingFolder_ContentPrepTool)"". Skipping download" -ForegroundColor Magenta
    }
    else {
        Get-FileFromInternet -URI "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" -Destination $WorkingFolder_ContentPrepTool
    }
} 
else {
    Write-Host "The -PackageApps parameter was not passed. Skipping downloading of the Win32 Content Prep Tool." -ForegroundColor Magenta
}
#EndRegion Get_Content_Tool


#Region Display_Application_Results
Write-Host ''
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Checking Applications...' -ForegroundColor DarkGray
Write-Host '--------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

#Get list of Applications
$ApplicationName = Get-CMApplication -Fast | Where-Object { $_.LocalizedDisplayName -like $AppName } | Select-Object -ExpandProperty LocalizedDisplayName | Sort-Object | Out-GridView -PassThru -OutVariable $ApplicationName

If ($ApplicationName) {
    Write-Host "The Win32App Migration Tool will proces the following Applications:"
    ForEach ($Application in $ApplicationName) {
        Write-Host """$($Application)""" -ForegroundColor Green
    }
}
#EndRegion Display_Application_Results

#Region Export_Details_CSV
#Call function to grab deployment type detail for application
$App_Array = Get-AppInfo -ApplicationName $ApplicationName
$DeploymentTypes_Array = $App_Array[0]
$Applications_Array = $App_Array[1]
$Content_Array = $App_Array[2]

#Export $DeploymentTypes to CSV for reference
Try {
    $DeploymentTypes_Array | Export-Csv (Join-Path -Path $WorkingFolder_Detail -ChildPath "DeploymentTypes.csv") -NoTypeInformation -Force
}
Catch {
    Write-Host "Error: Could not Export DeploymentTypes.csv. Do you have it open?" -ForegroundColor Red
}
Try {
    $Applications_Array | Export-Csv (Join-Path -Path $WorkingFolder_Detail -ChildPath "Applications.csv") -NoTypeInformation -Force
}
Catch {
    Write-Host "Error: Could not Export Applications.csv. Do you have it open?" -ForegroundColor Red
}
Try {
    $Content_Array | Export-Csv (Join-Path -Path $WorkingFolder_Detail -ChildPath "Content.csv") -NoTypeInformation -Force
}
Catch {
    Write-Host "Error: Could not Export DeploymentTypes.csv. Do you have it open?" -ForegroundColor Red
}
Write-Host "Details of the selected Applications and Deployment Types can be found at ""$($WorkingFolder_Detail)"""
#EndRegion Export_Details_CSV

#Region Exporting_Logos
If ($ExportLogo) {

    #Call function to export logo for application
    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Exporting Logo(s)...' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    ForEach ($Application in $Applications_Array) {
        $IconId = $Application.Application_IconId
        Export-Logo -IconId $IconId -AppName $Application.Application_Name
    }
}
#EndRegion Exporting_Logos

#Region Package_Apps
#If the $PackageApps parameter was passed. Use the Win32Content Prep Tool to build Intune.win files
If ($PackageApps) {
    #Region Creating_Application_Folders
    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Creating Application Folder(s)' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    ForEach ($Application in $Applications_Array) {

        #Create Application Parent Folder(s)
        Write-Host "Application: ""$($Application.Application_Name)"""
        Write-Host "Creating Application Folder ""$($Application.Application_LogicalName)"" for Application ""$($Application.Application_Name)""" -ForegroundColor Cyan
        If (!(Test-Path -Path (Join-Path -Path $WorkingFolder_Win32Apps -ChildPath $Application.Application_LogicalName ))) {
            New-FolderToCreate -Root $WorkingFolder_Win32Apps -Folders $Application.Application_LogicalName
        }
        else {
            Write-Host "Information: Application Folder ""$($Application.Application_LogicalName)"" already exists" -ForegroundColor Magenta
        }
        Write-Host ''
    }
    #EndRegion Creating_Application_Folders

    #Region Creating_DeploymentType_Folders
    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Creating DeploymentType Folder(s)' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''
    ForEach ($DeploymentType in $DeploymentTypes_Array) {

        #Create DeploymentType Child Folder(s)
        Write-Host "Creating DeploymentType Folder ""$($DeploymentType.DeploymentType_LogicalName)"" for DeploymentType ""$($DeploymentType.DeploymentType_Name)""" -ForegroundColor Cyan
        If (!(Test-Path -Path (Join-Path -Path (Join-Path -Path $WorkingFolder_Win32Apps -ChildPath $DeploymentType.Application_LogicalName ) -ChildPath $DeploymentType.DeploymentType_LogicalName))) {
            New-FolderToCreate -Root $WorkingFolder_Win32Apps -Folders (Join-Path -Path $DeploymentType.Application_LogicalName -ChildPath $DeploymentType.DeploymentType_LogicalName)
        }
        else {
            Write-Host "Information: Folder ""$($WorkingFolder_Win32Apps)\$($DeploymentType.DeploymentType_LogicalName)\$($DeploymentType.DeploymentType_LogicalName)"" already exists" -ForegroundColor Magenta
        }
        Write-Host ''
    }
    #EndRegion Creating_DeploymentType_Folders

    #Region Creating_Content_Folders
    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Creating Content Folder(s)' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''
    ForEach ($DeploymentType in $DeploymentTypes_Array) {

        #Create DeploymentType Content Folder(s)
        Write-Host "Creating DeploymentType Content Folder for DeploymentType ""$($DeploymentType.DeploymentType_Name)""" -ForegroundColor Cyan
        If (!(Test-Path -Path (Join-Path -Path $WorkingFolder_Content -ChildPath $DeploymentType.Application_LogicalName))) {
            New-FolderToCreate -Root $WorkingFolder_Content -Folders $DeploymentType.DeploymentType_LogicalName
        }
        else {
            Write-Host "Information: Folder ""$($WorkingFolder_Content)\$($DeploymentType.DeploymentType_LogicalName)"" Content already exists" -ForegroundColor Magenta
        }
        Write-Host ''
    }
    #EndRegion Creating_Content_Folders

    #Region Downloading_Content
    If ($PackageApps) {
        Write-Host '--------------------------------------------' -ForegroundColor DarkGray
        Write-Host 'Downloading Content' -ForegroundColor DarkGray
        Write-Host '--------------------------------------------' -ForegroundColor DarkGray
        Write-Host ''

        ForEach ($Content in $Content_Array) {
            Write-Host "Downloading Content for Deployment Type ""$($Content.Content_DeploymentType_LogicalName)"" from Content Source ""$($Content.Content_Location)""..." -ForegroundColor Cyan
            #$Files = Get-ChildItem -Path $Content.Content_Location -recurse | Select-Object -ExpandProperty Name
            Get-ContentFiles -Source $Content.Content_Location -Destination (Join-Path -Path $WorkingFolder_Content -ChildPath $Content.Content_DeploymentType_LogicalName)
        }
    }
    else {
        Write-Host "The -PackageApps parameter was not passed. Application and Deployment Type information will be gathered only, content will not be downloaded" -ForegroundColor Magenta
    }
    #EndRegion Downloading_Content

    #Region Create_Intunewin_Files
    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Creating .IntuneWin File(s)' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray

    #Get Application and Deployment Type Details and Files
    ForEach ($Application in $Applications_Array) {

        Write-Host ''
        Write-Host '--------------------------------------------' -ForegroundColor DarkGray
        Write-Host """$($Application.Application_Name)""" -ForegroundColor Green
        Write-Host "There are $($Application.Application_TotalDeploymentTypes) Deployment Types for this Application:"
        Write-Host '--------------------------------------------' -ForegroundColor DarkGray
        Write-Host ''

        ForEach ($Deployment in $DeploymentTypes_Array | Where-Object { $_.Application_LogicalName -eq $Application.Application_LogicalName }) {
            
            Write-Host '--------------------------------------------' -ForegroundColor DarkGray
            Write-Host """$($Deployment.DeploymentType_Name)""" -ForegroundColor Green
            Write-Host '--------------------------------------------' -ForegroundColor DarkGray
            Write-Host ''

            #Grab install command executable or script
            $SetupFile = $Deployment.DeploymentType_InstallCommandLine
            Write-Host "Install Command: ""$($SetupFile)"""

            ForEach ($Content in $Content_Array | Where-Object { $_.Content_DeploymentType_LogicalName -eq $Deployment.DeploymentType_LogicalName }) {

                #Create variables to pass to Function
                $ContentFolder = Join-Path -Path $WorkingFolder_Content -ChildPath $Deployment.DeploymentType_LogicalName
                $OutputFolder = Join-Path -Path (Join-Path -Path $WorkingFolder_Win32Apps -ChildPath $Application.Application_LogicalName) -ChildPath $Deployment.DeploymentType_LogicalName
                $SetupFile = $Deployment.DeploymentType_InstallCommandLine

                Write-Host "Content Folder: ""$($ContentFolder)"""
                Write-Host "Intunewin Output Folder: ""$($OutputFolder)"""
                Write-Host ''
                Write-Host "Creating .Intunewin for ""$($Deployment.DeploymentType_Name)""..." -ForegroundColor Cyan
                $IntuneWinFileCommand = New-IntuneWin -ContentFolder $ContentFolder -OutputFolder $OutputFolder -SetupFile $SetupFile
                $IntuneWinFile = $IntuneWinFileCommand
                Write-Host ''
                 
                If (Test-Path (Join-Path -Path $OutputFolder -ChildPath "*.intunewin") ) {
                    Write-Host "Successfully created ""$($IntuneWinFile).intunewin"" at ""$($OutputFolder)""" -ForegroundColor Cyan
                }
                else {
                    Write-Host "Error: We couldn't verify that ""$($IntuneWinFile).intunewin"" was created at ""$($OutputFolder)""" -ForegroundColor Red
                }
            }
        }
    }
}
#EndRegion Create_Intunewin_Files

#EndRegion Package_Apps

#Region Create_Apps
#If the $CreateApps parameter was passed. Use the Win32Content Prep Tool to create Win32 Apps
If ($CreateApps) {

    Write-Host ''
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Creating Win32 Apps' -ForegroundColor DarkGray
    Write-Host '--------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''
}
#EndRegion Create_Apps
Set-Location $ScriptRoot
Write-Host ''
Write-Host '## The Win32AppMigrationTool Script has Finished ##'