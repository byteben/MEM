???<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.188
	 Created on:   	31/05/2021 10:01
	 Created by:   	MauriceDaly
	 Organization: 	CloudWay
	 Filename:     	Invoke-TenantSecurity.ps1
	===========================================================================
	.DESCRIPTION
		Sets recommendations from the Microsoft Security Center
#>

# Enable 'Local Security Authority (LSA) protection
New-ItemProperty -Path HKLM:\System\CurrentControlSet\Control\Lsa -Name "RunAsPPL" -Value 1 -PropertyType DWORD -Force

#Setting registry key to block AAD Registration to 3rd party tenants. 
$RegistryLocation = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin\"
$keyname = "BlockAADWorkplaceJoin"

#Test if path exists and create if missing
if (!(Test-Path -Path $RegistryLocation)) {
	Write-Output "Registry location missing. Creating"
	New-Item $RegistryLocation | Out-Null
}

#Force create key with value 1 
New-ItemProperty -Path $RegistryLocation -Name $keyname -PropertyType DWord -Value 1 -Force | Out-Null
Write-Output "Registry key set"

# Adobe Acrobat 
$AdobeRegKey = "HKLM:\SOFTWARE\Policies\Adobe"
if (Test-Path -Path $AdobeRegKey -eq $true) {
	if (Test-Path -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -eq $true) {
		New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bDisableJavaScript" -PropertyType DWord -Value 1 -Force | Out-Null
		New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bEnableFlash" -PropertyType DWord -Value 0 -Force | Out-Null
	}
	if (Test-Path -Path "HKLM\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown" -eq $true) {
		New-ItemProperty -Path "HKLM\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown" -Name "bDisableJavaScript" -PropertyType DWord -Value 0 -Force | Out-Null
		
	}
}



