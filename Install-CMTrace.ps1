<#
.SYNOPSIS
    Download and Install the Microsoft ConfigMgr 2012 Toolkit R2

.DESCRIPTION
    This script will download and install the Microsoft ConfigMgr 2012 Toolkit R2. 
    The intention was to use the scrtipt for Intune managed devices to easily read log files in the absence of the ConfigMgr client and C:\Windows\CCM\CMTrace.exe

.EXAMPLE
    .\Install-CMTrace.exe

.NOTES
    FileName:    Install-CMTrace.ps1
    Author:      Ben Whitmore
    Date:        9th July 2022
    Thanks:      @PMPC colleague(s) for "Current user" identitity and role functions
    
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$URL = "https://download.microsoft.com/download/5/0/8/508918E1-3627-4383-B7D8-AA07B3490D21/ConfigMgrTools.msi",
    [Parameter(Mandatory = $false)]
    [string]$DownloadDir = $env:temp,
    [Parameter(Mandatory = $false)]
    [string]$DownloadFileName = "ConfigMgrTools.msi",
    [Parameter(Mandatory = $false)]
    [string]$InstallDest = "C:\Program Files (x86)\ConfigMgr 2012 Toolkit R2\ClientTools\CMTrace.exe"
)

##Set Verbose Level##
$VerbosePreference = "Continue"
#$VerbosePreference = "SilentlyContinue"

Function Get-URLHashInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$URLPath
    )
    #Initialize WebClient class and return file hash of URI passed to the function
    $WebClient = [System.Net.WebClient]::new()
    $URLHash = Get-FileHash -Algorithm MD5 -InputStream ($WebClient.OpenRead($URLPath))
    return $URLHash
}

Function Get-FileHashInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    #Return file hash of file passed to the function
    $FileHash = Get-FileHash -Algorithm MD5 -Path $FilePath
    return $FileHash
}

Function Get-CurrentUser {

    #Get the current user
    $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    Write-Verbose "Current User = $($CurrentUser.Name)"
    return $CurrentUser
}

Function Test-IsRunningAsAdministrator {

    #Get current user and return $true if they are a local administrator
    $CurrentUser = Get-CurrentUser
    $IsAdmin = (New-Object Security.Principal.WindowsPrincipal $CurrentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    Write-Verbose "Current user is an administrator? $IsAdmin"
    return $IsAdmin
}

Function Test-IsRunningAsSystem {

    #Get current user and return $true if it is SYSTEM
    $RunningAsSystem = (Get-CurrentUser).User -eq 'S-1-5-18'
    Write-Verbose "Running as system? $RunningAsSystem"
    return $RunningAsSystem
}

Function Get-FileFromInternet {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_.StartsWith("https://") })]
        [String]$URL,
        [Parameter(Mandatory = $true)]
        [String]$Destination
    )

    #Test the URL is valid
    Try {
        $URLRequest = Invoke-WebRequest -UseBasicParsing -URI $URL -ErrorAction SilentlyContinue
        $StatusCode = $URLRequest.StatusCode
    }
    Catch {
        Write-Verbose "It looks like the URL is invalid. Please try again"
        $StatusCode = $_.Exception.Response.StatusCode.value__
    }

    #If the URL is valid, attempt to download the file
    If ($StatusCode -eq 200) {
        Try {
            Invoke-WebRequest -UseBasicParsing -Uri $URL -OutFile $Destination -ErrorAction SilentlyContinue

            If (Test-Path $Destination) {
                Write-Verbose "File download Successfull. File saved to $Destination"
            }
            else {
                Write-Verbose "The download was interrupted or an error occured moving the file to the destination you specified"
                break
            }
        }
        Catch {
            Write-Verbose "Error downloading file: $URL"
            $_
            break
        }
    }
    else {
        Write-Verbose "URL Does not exists or the website is down. Status Code: $($StatusCode)"
        break
    }
}

#Continue if the user running the script is a local administrator or SYSTEM
If ((!(Test-IsRunningAsAdministrator)) -and (!(Test-IsRunningAsSystem))) {
    Write-Verbose "The current User is not an administrator or SYSTEM. Please run this script with administrator credentials or in the SYSTEM context"
}
else {

    #Check the CLient Tools are not already installed
    If (-not(Test-Path -Path $InstallDest)) {

        #Set Download destination full path
        $FilePath = (Join-Path -Path $DownloadDir -ChildPath $DownloadFileName)

        #Download the MSI
        Get-FileFromInternet -URL $URL -Destination $FilePath

        #Test if the download was successful
        If (-not (Test-Path -Path $FilePath)) {
            Write-Verbose "There was an error downloading the file to the file system. $($FilePath) does not exist."
            break
        }
        else {
            
            #Check the hash from Microsoft matches the hash of the file saved to disk
            $URLHash = (Get-URLHashInfo -URLPath $URL).hash
            $FileHash = (Get-FileHashInfo -FilePath $FilePath).hash
            Write-Verbose "Checking Hash.."

            #Warn if the hash is different
            If (($URLHash -ne $FileHash) -or ([string]::IsNullOrWhitespace($URLHash)) -or ([string]::IsNullOrWhitespace($FileHash))) {
                Write-Verbose "URL Hash = $($URLHash)"
                Write-Verbose "File Hash = $($FileHash)"
                Write-Verbose "There was an error checking the hash value of the downloaded file. The file hash for ""$($FilePath)"" does not match the hash at ""$($URL)"". Aborting installation"
                break
            }
            else {
                Write-Verbose "Hash match confirmed. Continue installation.."
                Try {

                    #Attempt to install the MSI
                    $MSIArgs = @(
                        "/i"
                        $FilePath
                        "ADDLOCAL=ClientTools"
                        "/qn"
                    )
                    Start-Process "$env:SystemRoot\System32\msiexec.exe" -args $MSIArgs -Wait -NoNewWindow

                    #Check the installation was successful
                    If (Test-Path -Path $InstallDest) {
                        Write-Verbose "CMTrace installed succesfully at $($InstallDest) "
                    }
                }
                Catch {
                    Write-Verbose "There was an error installing the CMTrace"
                    $_
                }
            }
        }
    }
    else {
        Write-Verbose "CMTrace is already installed at $($InstallDest). Installation will not continue."
    }
}