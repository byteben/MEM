<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.188
	 Created on:   	24/05/2021 15:46
	 Created by:   	MauriceDaly
	 Organization: 	CloudWay
	 Filename:     	Remove-AppxUserPackages.ps1
	===========================================================================
	.DESCRIPTION
		Removes unwanted Appx packages
#>

$BlackListedUserApps = New-Object -TypeName System.Collections.ArrayList
$BlackListedUserApps.AddRange(@(
		"Microsoft.BingWeather",
		"Microsoft.Xbox",
		"Microsoft.People",
		"Microsoft.Microsoft3DViewer",
		"Microsoft.OfficeHub",
		"Microsoft.SolitaireCollection",
		"Microsoft.Microsoft3DViewer",
		"Microsoft.SkypeApp",
		"Microsoft.WindowsCommunicationsApps"
	))

$AppArrayList = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName
foreach ($App in $BlackListedUserApps) {
	if (($AppArrayList -match $App)) {
		Get-AppxPackage -AllUsers | Where-Object {$_.Name -match $App } | Remove-AppxPackage -ErrorAction SilentlyContinue
	}
}





