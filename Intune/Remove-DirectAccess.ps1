<#
.Synopsis
Created on:   10/08/2021
Created by:   Ben Whitmore @ CloudWay
Filename:     Remove-DirectAccess.ps1

.Description
Script to remove Direct Access client settings
#>

$Path = 'HKLM:\Software\Policies\Microsoft\Windows NT\DNSClient\DnsPolicyConfig'
If (Test-Path $Path) {
    Write-Output "Direct Access DnsPolicyConfig registry key found.."
    Try {
        Write-Output "Attempting to delete Direct Access DnsPolicyConfig registry key.."
        Get-Item -Path $Path | Remove-Item -Recurse -Confirm:$False
        If (Test-Path $Path) {
            Write-Output "There was a problem removing the registry key"
            Exit 1 
        }
        else {
            Write-Output "Direct Access DnsPolicyConfig registry key deleted successfully"
            Exit 0
        }
    }
    Catch {
        Write-Output "There was a problem removing the registry key"
        Exit 1
    }
}
else {
    Write-Output "Direct Access DnsPolicyConfig registry key not found"
    Exit 0
}