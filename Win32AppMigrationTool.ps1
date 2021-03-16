﻿Param (
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

Function Connect-SiteServer {
    Param (
        [String]$SiteCode,
        [String]$ProviderMachineName
    )
    # Import the ConfigurationManager.psd1 module 
    if ($Null -eq (Get-Module ConfigurationManager)) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    }

    # Connect to the site's drive if it is not already present
    if ($Null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
    }

    #Set the current location to be the site code.
    Set-Location "$($SiteCode):\"
}

#Create Global Variables
$Global:WorkingFolder_Root = $WorkingFolder
$Global:WorkingFolder_Logos = Join-Path -Path $WorkingFolder_Root -ChildPath "Logos"
$Global:WorkingFolder_ContentPrepTool = Join-Path -Path $WorkingFolder_Root -ChildPath "ContentPrepTool"
$Global:WorkingFolder_Logs = Join-Path -Path $WorkingFolder_Root -ChildPath "Logs"
$Global:WorkingFolder_Detail = Join-Path -Path $WorkingFolder_Root -ChildPath "Details"
$Global:WorkingFolder_Win32Apps = Join-Path -Path $WorkingFolder_Root -ChildPath "Win32Apps"
Function New-FolderToCreate {
    <#
    Function to create folder structure for Win32AppMigrationTool
    #>  
    Param(
        [String]$Root,
        [String[]]$Folders
    )
    If (!($Root)) {
        Write-Warning "No Root Folder passed to Function"
    }
    If (!($Folders)) {
        Write-Warning "No Folder(s) passed to Function"
    }

    ForEach ($Folder in $Folders) {
        #Create Folders
        $FolderToCreate = Join-Path -Path $Root -ChildPath $Folder
        If (!(Test-Path $FolderToCreate)) {
            Write-Output "Creating Folder ""$($FolderToCreate)"""
            Try {
                New-Item -Path $FolderToCreate -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Output "Folder ""$($FolderToCreate)"" created succesfully"
            }
            Catch {
                Write-Warning "Couldn't create ""$($FolderToCreate)"" folder"
            }
        }
        else {
            Write-Output "Folder ""$($FolderToCreate)"" already exsts. Skipping.."
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
    Write-Output "Preparing to export Application Logo for ""$($AppName)"""
    If ($IconId) {

        #Check destination folder exists for logo
        If (!(Test-Path $WorkingFolder_Logos)) {
            Try {
                New-Item -Path $WorkingFolder_Logos -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            Catch {
                Write-Warning "Couldn't create ""$($WorkingFolder_Logos)"" folder for Application Logos"
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
                        Write-Warning "Couldn't create ""$($LogoFolder_Id)"" folder for Application Logo "
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
                            Write-Output "Application logo for ""$($AppName)"" exported successfully to ""$($Logo_File)"""
                        }
                    }
                    Catch {
                        Write-Warning "Could not export Logo to folder ""$($LogoFolder_Id)"""
                    }
                }
            }
            else {
                Write-Warning "Did not export Logo for ""$($AppName)"" to ""$($Logo_File)"" because the file already exists"
            }
        }
    }
    else {
        Write-Warning "Null or invalid IconId passed to function. Could not export Logo"
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
    $file
    Try {
        Invoke-WebRequest -UseBasicParsing -Uri $URI -OutFile $FileDestination -ErrorAction Stop
    }
    Catch {
        Write-Warning "Error downloading the Win32 Content Prep Tool"
        $_
    }
}
#Connect to Site Server
Connect-SiteServer -SiteCode  $SiteCode -ProviderMachineName $ProviderMachineName

#Create Folders
Write-Output "Setting up Environment for Win32App Migration Tool"
Write-Output "Creating Folders..."
New-FolderToCreate -Root $WorkingFolder_Root -Folders @("", "Logos", "ContentPrepTool", "Logs", "Details", "Win32Apps")

#Download Win32 Content Prep Tool
Write-Output "Downloadling Win32 Content Prep Tool..."
If (Test-Path (Join-Path -Path $WorkingFolder_ContentPrepTool -ChildPath "IntuneWinAppUtil.exe")) {
    Write-Output "IntuneWinAppUtil.exe already exists at ""$($WorkingFolder_ContentPrepTool)"". Skipping download"
}
else {
    Write-Output "Downloading Win32 Content Prep Tool..."
    Get-FileFromInternet -URI "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/blob/master/IntuneWinAppUtil.exe" -Destination $WorkingFolder_ContentPrepTool
}

#Get list of Applications
$ApplicationName = Get-CMApplication -Fast | Where-Object { $_.LocalizedDisplayName -like $AppName } | Select-Object -ExpandProperty LocalizedDisplayName | Sort-Object

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
                    $ApplicationTypes | Add-Member NoteProperty -Name Application_IconPath -Value (Join-Path -Path $WorkingFolder_Logos -ChildPath (Join-Path -Path $XMLContent.AppMgmtDigest.Application.DisplayInfo.Info.Icon.Id -ChildPath "Logo.jpg"))
                }
                else {
                    $ApplicationTypes | Add-Member NoteProperty -Name Application_IconPath -Value $Null
                }
                $ApplicationTypes += $ApplicationObject
        
            #If Deployment Types exist, iterate through each DeploymentType and build deployment detail
            ForEach ($Object in $XMLContent.AppMgmtDigest.DeploymentType) {

                #Create new custom PSObject to build line detail
                $DeploymentObject = New-Object PSCustomObject
                
                #DeploymentType Details
                $DeploymentObject | Add-Member NoteProperty -Name Application_LogicalName -Value $XMLContent.AppMgmtDigest.Application.LogicalName
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
            }
        }
    } 
    Return $DeploymentTypes, $ApplicationTypes
}

If ($ApplicationName) {
    Write-Output "Found the following matches for ""$($AppName)"""
    ForEach ($Application in $ApplicationName) {
        Write-Output """$($Application)"""
    }
}

#Call function to grab deployment type detail for application
$App_Array = Get-AppInfo -ApplicationName $ApplicationName
$DeploymentTypes_Array = $App_Array[0]
$Applications_Array = $App_Array[1]

#Export $DeploymentTypes to CSV for reference
$DeploymentTypes_Array | Export-Csv (Join-Path -Path $WorkingFolder_Detail -ChildPath "DeploymentTypes.csv") -NoTypeInformation -Force
$Applications_Array | Export-Csv (Join-Path -Path $WorkingFolder_Detail -ChildPath "Applications.csv") -NoTypeInformation -Force

#Call function to export logo for application
If ($ExportLogo) {
    ForEach ($Application in $Applications_Array) {
        $IconId = $Application.Application_IconId
        Export-Logo -IconId $IconId -AppName $Application.Application_Name
    }
}

#If the $PackageApps parameter was passed. Use the Win32Content Prep Tool to build Intune.win files
If ($PackageApps) {
    Write-Output "Preparing to create Intune.win files..."
    ForEach ($Application in $Applications_Array) {

        #Create Folders
        Write-Output "Creating Folder ""$($Application.Application_LogicalName)"" for Application ""$($Application.Application_Name)"""
        If (!(Test-Path -Path (Join-Path -Path $WorkingFolder_Win32Apps -ChildPath $Application.Application_LogicalName ))) {
            New-FolderToCreate -Root $WorkingFolder_Win32Apps -Folders $Application.Application_LogicalName
        }
        else {
            Write-Warning "Folder ""$($Application.Application_LogicalName)"" Exists" 
        }
    }
}