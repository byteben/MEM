<#	
	.NOTES
	===========================================================================
	 Created on:   	21/8/2020 11:00 PM
	 Created by:   	Maurice Daly
	 Organization: 	CloudWay
	 Filename:     	Invoke-BitLockerCompliance.ps1
	===========================================================================
	.DESCRIPTION
		Enforce BitLocker compliance through detection and remediation functions
#>

function Write-LogEntry {
	param (
		[parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("1", "2", "3")]
		[string]$Severity,
		[parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = "Invoke-BitLockerCompliance.log"
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $env:WinDir -ChildPath "Temp\$FileName"
	
	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")
	
	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
	
	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""BitLocker Recovery Key Backup"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
		if ($Severity -eq 1) {
			Write-Verbose -Message $Value
		} elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	} catch [System.Exception] {
		Write-Warning -Message "Unable to append log entry to InvokeBitLockerBackup.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}

function Enforce-BitLocker {
	try {
		
		$ComputerSystemType = Get-WmiObject -Class "Win32_ComputerSystem" | Select-Object -ExpandProperty "Model"
		if ($ComputerSystemType -notin @("Virtual Machine", "VMware Virtual Platform", "VirtualBox", "HVM domU", "KVM", "VMWare7,1")) {
			
			# Check if machine is a member of a domain
			Write-LogEntry "Checking domain membership" -Severity 1
			$DomainMember = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty PartOfDomain
			
			# Obtain BitLocker information
			Write-LogEntry "Obtaining BitLocker volume information for C:\" -Severity 1
			$BitLockerVolume = Get-BitLockerVolume -MountPoint C: | Select-Object -Property *
			
			# Obtain TPM information
			$TPMValues = Get-Tpm -ErrorAction SilentlyContinue | Select-Object -Property TPMReady, TPMPresent, TPMEnabled, TPMActivated, TPMOwned, ManagedAuthLevel
			
			if ($BitLockerVolume.VolumeStatus -match "encrypt" -and $BitLockerVolume.ProtectionStatus -eq "On") {
				Write-LogEntry "Drive encryption is in place. Starting backup process." -Severity 1
				
				# Set current protector to backup
				$CurrentProtector = 0
				$TotalProtectorKeys = $($BitLockerVolume.KeyProtector | Where-Object {
						$_.RecoveryPassword -gt $null
					}).Count
				$KeyProtectors = $($BitLockerVolume.KeyProtector | Where-Object {
						(-not ([string]::IsNullOrEmpty($_.RecoveryPassword)))
					})
				Write-LogEntry "Backing up $TotalProtectorKeys recovey keys" -Severity 1
				
				foreach ($KeyProtector in $KeyProtectors) {
					if ($DomainMember -eq $true) {
						# Backup to Active Directory
						Write-LogEntry "Backing up recovery key $CurrentProtector to Active Directory" -Severity 1
						Write-LogEntry "Key protector ID is $($KeyProtector.KeyProtectorID)" -Severity 1
						Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $KeyProtector.KeyProtectorID -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
					}
					# Backup to Azure Active Directory
					Write-LogEntry "Backing up recovery key $CurrentProtector to Azure Active Directory" -Severity 1
					Write-LogEntry "Key protector ID is $($KeyProtector.KeyProtectorID)" -Severity 1
					BackupToAAD-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $KeyProtector.KeyProtectorID | Out-Null
					$CurrentProtector++
				}
				
				Write-LogEntry "Recovery key backup process complete." -Severity 1
				exit 0
				
			} else {
				Write-LogEntry "Drive is currently not protected by BitLocker. Starting remediation process.." -Severity 2
				if ($BitLockerVolume.EncryptionPercentage -ne '100' -and $BitLockerVolume.EncryptionPercentage -ne '0') {
					Write-LogEntry "BitLocked detected, drive is not fully encrypted. Finishing encryption process" -Severity 2
					Write-Output "BitLocked detected, drive is not fully encrypted. Finishing encryption process"
					Resume-BitLocker -MountPoint "C:" | Out-Null
					$ErrorMessage = "Failed to resume BitLocker protection"
					
				} elseif ($BitLockerVolume.VolumeStatus -eq 'FullyEncrypted' -and $BitLockerVolume.ProtectionStatus -eq 'Off') {
					Write-LogEntry "BitLocked detected, but disabled. Re-enabling BitLocker protection." -Severity 2
					Write-Output "BitLocked detected, but disabled. Re-enabling BitLocker protection."
					Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAES256 -UsedSpaceOnly -SkipHardwareTest -TpmProtector | Out-Null
					Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAES256 -UsedSpaceOnly -SkipHardwareTest -RecoveryPasswordProtector | Out-Null
					Resume-BitLocker -MountPoint "C:" | Out-Null
					$ErrorMessage = "Failed to resume BitLocker protection"
					
				} elseif ($BitLockerVolume.VolumeStatus -eq 'FullyDecrypted' -and $TPMValues.TPMEnabled -eq $true) {
					Write-LogEntry "BitLocked encryption not present. Setting to enabled with XTSAES256 encryption cipher specified." -Severity 2
					Write-Output "BitLocked encryption not present. Setting to enabled with XTSAES256 encryption cipher specified."
					Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAES256 -UsedSpaceOnly -SkipHardwareTest -TpmProtector | Out-Null
					Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAES256 -UsedSpaceOnly -SkipHardwareTest -RecoveryPasswordProtector | Out-Null
					$ErrorMessage = "Failed to enable BitLocker to encrypt drive using TPM protector"
				}
				
				Start-Sleep -Seconds 10
				$BitLockerVolume = Get-BitLockerVolume -MountPoint "C:" | Select-Object *
				if ($BitLockerVolume.ProtectionStatus -eq "On") {
					Write-LogEntry "BitLocker enabled on drive C:" -Severity 1
					Write-LogEntry "Backing up BitLocker encryption key" -Severity 1
					BackupToAAD-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $BitLockerVolume.KeyProtector[1].KeyProtectorId
					# Refresh BitLocker drive information
					Write-LogEntry "BitLocker encryption method : $($BitLockerVolume.EncryptionMethod)" -Severity 1
					Write-LogEntry "BitLocker volume status : $($BitLockerVolume.VolumeStatus)" -Severity 1
					Write-LogEntry "BitLocker key protector types : $($BitLockerVolume.KeyProtector)" -Severity 1
					Write-LogEntry "Drive capacity : $($BitLockerVolume.CapacityGB)GB" -Severity 1
					Write-Output "BitLocker protection status OK"; exit 0
				} else {
					Write-LogEntry "$ErrorMessage" -Severity 3
					Write-Output "$ErrorMessage"; exit 1
				}
			}
		} else {
			Write-LogEntry "Virtual machine detected. Skipping BitLocker enforcement" -Severity 2
			Write-Output "Virtual machine detected. Skipping BitLocker enforcement"; exit 0
		}
	} catch {
		Write-LogEntry -Value "Issues occured during the key recovery backup process $($_.Exception.Message)" -Severity 3
		Write-Output "Issues occured during the key recovery backup process $($_.Exception.Message)"; exit 1
	}
}
Enforce-BitLocker

	
