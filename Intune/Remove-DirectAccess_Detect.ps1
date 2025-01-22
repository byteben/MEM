<#
.Synopsis
Created on:   10/08/2021
Created by:   Ben Whitmore @ CloudWay
Filename:     Remove-DirectAccess_Detect.ps1

.Description
Script to remove Direct Access client settings
#>

$Path = 'HKLM:\Software\Policies\Microsoft\Windows NT\DNSClient\DnsPolicyConfig'
If (Test-Path $Path) {
    Write-Output "Direct Access DnsPolicyConfig registry key found. Return 1"
    #Exit 1
}
else {
    Write-Output "Direct Access DnsPolicyConfig registry key not found. Return 0"
    #Exit 0
}