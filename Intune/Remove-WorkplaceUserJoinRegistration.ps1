Param(
    [Parameter(Mandatory = $False)]
    [Switch]$Remove,
    [Parameter(Mandatory = $False)]
    [String[]] $Domains = @("tawkr.agency", "appcogroupnederland.com")
)

Function Get-WorkplaceJoinUserReg {
    $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo"

    Try {
        $RegKeys = (Get-ChildItem $RegPath).PSChildName
    }
    Catch {
        Write-Warning "Unable to get corresponding Registry Keys for Workplace Join User Registrations"
    }

    If ($RegKeys) {
        $Keys = @()
        Foreach ($Reg in $Regkeys) {
            $Tenant = $Null
            Try {
                $WorkPlaceReg = Join-Path -Path $RegPath -Childpath $Reg 
                Foreach ($Domain in $Domains) {  
                    $Tenant = Get-ItemProperty $WorkPlaceReg | Select-Object -ExpandProperty Useremail | Where-Object { $_ -like "*$Domain" } -ErrorAction SilentlyContinue
                    If ($tenant) {
                        If (!($Remove)) {
                            Write-Host "Workplace Join User Registration Registry key found for $Tenant and ID $Reg"
                        }
                        $Keys += $Reg 
                        
                    } 
                }
            }
            Catch {
                Write-Warning "Unable to get corresponding Registry Keys for Workplace Join User Registrations"
            }
        }
    }
    
    Return $Keys
    
}

Function Get-WorkplaceJoinUserCert {
    $Registrations = Get-WorkplaceJoinUserReg
    $Certs = @()
    Foreach ($Registration in $Registrations) {
        Try {
            $AAD_Cert = Get-ChildItem -path Cert:\CurrentUser\My | Where-Object { $_.ThumbPrint -eq $Registration } -ErrorAction SilentlyContinue
            If ($AAD_Cert) {
                If (!($Remove)) {
                    Write-Host "Found AAD Workplace Join User Certificate with matching ThumbPrint ""$($Registration)"""
                }
                $Certs += $Registration
            }
            else {
                Write-Warning "No AAD Workplace Join User Certificate found with matching ThumbPrint ""$($Registration)"""
            }
        }
        Catch {
            Write-Warning "Unable to find Certificate with Thumbprint ""$($Registration)"""
        }
    }

    Return $certs
}

Function Get-WorkplaceJoinUserInfo {
    $Info = Get-WorkplaceJoinUserCert
}

Function Remove-WorkplaceJoinUserRegistration {

    $CertsToRemove = Get-WorkplaceJoinUserCert
    Foreach ($CertToRemove in $CertsToRemove) {
        Try {
            $AAD_Cert = Get-ChildItem -path Cert:\CurrentUser\My | Where-Object { $_.ThumbPrint -eq $Registration }
            Write-Host "Deleting AAD Workplace Join User Certificate with matching ThumbPrint ""$($Registration)"""
            $AAD_Cert | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Write-Host "Done"
        }
        Catch {
            Write-Warning "Unable to find Certificate with Thumbprint ""$($Registration)"""
        }
    }

    $RegsToRemove = Get-WorkplaceJoinUserReg

    $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo"

    Foreach ($RegToRemove in $RegsToRemove) {
        Try {
            $WorkPlaceReg = Join-Path -Path $RegPath -Childpath $RegToRemove
            Write-Host "Deleting AAD Workplace Join User Registry Key ""$($WorkPlaceReg)"""
            Remove-Item -Path $WorkPlaceReg -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Done"
        }
        Catch {
            Write-Warning "Error Removing Registry Key ""$WorkPlaceReg"""

        }
    }
}

If ($Remove) {
    Remove-WorkplaceJoinUserRegistration
}
else {
    Get-WorkplaceJoinUserInfo
}