<#
.SYNOPSIS
    Get Win32app state from the local registry

.DESCRIPTION
    This script will get the state of Win32apps from the local registry
    It will return the compliance state and enforcement state of the app, as well as the state results for each

    ##Example using the returned $stateMessages array to find apps that are applicable, required but not installed
    $stateMessages | where-Object {$_.complianceStateMessage.Applicability -eq 'Applicable' -and $_.complianceStateMessage.ComplianceState -eq 'NotInstalled' -and $_.complianceStateMessage.DesiredState -eq 'Present'}

.NOTES
    FileName:    Get-Win32AppResults.ps1
    Author:      Ben Whitmore
    Date:        28th August 2023

.PARAMETER Win32appKey
        The registry key to look for Win32apps under. Default is HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps

.PARAMETER DontTranslateStates
        If set, the script will not attempt to translate the numerical values of the state messages to the string equivilent

.EXAMPLE
    .\Get-Win32AppResults.ps1

.EXAMPLE
    .\Get-Win32AppResults.ps1 -DontTranslateStates
    
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Win32appKey = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps",
    [switch]$DontTranslateStates
)

# declare state message compliance state message hashtables
$stateMessageEnforcement = @{
    1000 = "Success"
    1003 = "SuccessFastNotify"
    1004 = "SuccessButDependencyFailedToInstall"
    1005 = "SuccessButDependencyWithRequirementsNotMet"
    1006 = "SuccessButDependencyPendingReboot"
    1007 = "SuccessButDependencyWithAutoInstallOff"
    1008 = "SuccessButIOSAppStoreUpdateFailedToInstall"
    1009 = "SuccessVPPAppHasUpdateAvailable"
    1010 = "SuccessButUserRejectedUpdate"
    1011 = "SuccessUninstallPendingReboot"
    1012 = "SuccessSupersededAppUninstallFailed"
    1013 = "SuccessSupersededAppUninstallPendingReboot"
    1014 = "SuccessSupersedingAppsDetected"
    1015 = "SuccessSupersededAppsDetected"
    1016 = "SuccessAppRemovedBySupersedence"
    1017 = "SuccessButDependencyBlockedByManagedInstallerPolicy"
    1018 = "SuccessUninstallingSupersededApps"
    2000 = "InProgress"
    2007 = "InProgressDependencyInstalling"
    2008 = "InProgressPendingReboot"
    2009 = "InProgressDownloadCompleted"
    2010 = "InProgressPendingUninstallOfSupersededApps"
    2011 = "InProgressUninstallPendingReboot"
    2012 = "InProgressPendingManagedInstaller"
    3000 = "RequirementsNotMet"
    4000 = "Unknown"
    5000 = "Error"
    5003 = "ErrorDownloadingContent"
    5006 = "ErrorConflictsPreventInstallation"
    5015 = "ErrorManagedInstallerAppLockerPolicyNotApplied"
    5999 = "ErrorWithImmeadiateRetry"
    6000 = "NotAttempted"
    6001 = "NotAttemptedDependencyWithFailure"
    6002 = "NotAttemptedPendingReboot"
    6003 = "NotAttemptedDependencyWithRequirementsNotMet"
    6004 = "NotAttemptedAutoInstallOff"
    6005 = "NotAttemptedDependencyWithAutoInstallOff"
    6006 = "NotAttemptedWithManagedAppNoLongerPresent"
    6007 = "NotAttemptedBecauseUserRejectedInstall"
    6008 = "NotAttemptedBecauseUserIsNotLoggedIntoAppStore"
    6009 = "NotAttemptedSupersededAppUninstallFailed"
    6010 = "NotAttemptedSupersededAppUninstallPendingReboot"
    6011 = "NotAttemptedUntargetedSupersedingAppsDetected"
    6012 = "NotAttemptedDependencyBlockedByManagedInstallerPolicy"
    6013 = "NotAttemptedUnsupportedOrIndeterminateSupersededApp"
}

$stateMessageApplicability = @{
    0    = "Applicable"
    1    = "RequirementsNotMet"
    3    = "HostPlatformNotApplicable"
    1000 = "ProcessorArchitectureNotApplicable"
    1001 = "MinimumDiskSpaceNotMet"
    1002 = "MinimumOSVersionNotMet"
    1003 = "MinimumPhysicalMemoryNotMet"
    1004 = "MinimumLogicalProcessorCountNotMet"
    1005 = "MinimumCPUSpeedNotMet"
    1006 = "FileSystemRequirementRuleNotMet"
    1007 = "RegistryRequirementRuleNotMet"
    1008 = "ScriptRequirementRuleNotMet"
    1009 = "NotTargetedAndSupersedingAppsNotApplicable"
    1010 = "AssignmentFiltersCriteriaNotMet"
    1011 = "AppUnsupportedDueToUnknownReason"
    1012 = "UserContextAppNotSupportedDuringDeviceOnlyCheckin"
    2000 = "COSUMinimumApiLevelNotMet"
    2001 = "COSUManagementMode"
    2002 = "COSUUnsupported"
    2003 = "COSUAppIncompatible"
}

$stateMessageComplianceState = @{
    1   = "Installed"
    2   = "NotInstalled"
    4   = "Error"
    5   = "Unknown"
    100 = "Cleanup"
}

$stateMessageDesiredState = @{
    0 = "None"
    1 = "Not Present"
    2 = "Present"
    3 = "Unknown"
    4 = "Available"
}

$stateMessageTargetingMethod = @{
    0 = "EgatTargetedApplication"
    1 = "DependencyOfEgatTargetedApplication"
}

$stateMessageInstallContext = @{
    1 = "User"
    2 = "System"
}

$stateMessageTargetType = @{
    0 = "None"
    1 = "User"
    2 = "Device"
    3 = "Both Device and User"
}

# declare state message enforcement state message hashtables
$stateMessageEnforcementState = @{
    1000 = "Success"
    1003 = "SuccessFastNotify"
    1004 = "SuccessButDependencyFailedToInstall"
    1005 = "SuccessButDependencyWithRequirementsNotMet"
    1006 = "SuccessButDependencyPendingReboot"
    1007 = "SuccessButDependencyWithAutoInstallOff"
    1008 = "SuccessButIOSAppStoreUpdateFailedToInstall"
    1009 = "SuccessVPPAppHasUpdateAvailable"
    1010 = "SuccessButUserRejectedUpdate"
    1011 = "SuccessUninstallPendingReboot"
    1012 = "SuccessSupersededAppUninstallFailed"
    1013 = "SuccessSupersededAppUninstallPendingReboot"
    1014 = "SuccessSupersedingAppsDetected"
    1015 = "SuccessSupersededAppsDetected"
    1016 = "SuccessAppRemovedBySupersedence"
    1017 = "SuccessButDependencyBlockedByManagedInstallerPolicy"
    1018 = "SuccessUninstallingSupersededApps"
    2000 = "InProgress"
    2007 = "InProgressDependencyInstalling"
    2008 = "InProgressPendingReboot"
    2009 = "InProgressDownloadCompleted"
    2010 = "InProgressPendingUninstallOfSupersededApps"
    2011 = "InProgressUninstallPendingReboot"
    2012 = "InProgressPendingManagedInstaller"
    3000 = "RequirementsNotMet"
    4000 = "Unknown"
    5000 = "Error"
    5003 = "ErrorDownloadingContent"
    5006 = "ErrorConflictsPreventInstallation"
    5015 = "ErrorManagedInstallerAppLockerPolicyNotApplied"
    5999 = "ErrorWithImmeadiateRetry"
    6000 = "NotAttempted"
    6001 = "NotAttemptedDependencyWithFailure"
    6002 = "NotAttemptedPendingReboot"
    6003 = "NotAttemptedDependencyWithRequirementsNotMet"
    6004 = "NotAttemptedAutoInstallOff"
    6005 = "NotAttemptedDependencyWithAutoInstallOff"
    6006 = "NotAttemptedWithManagedAppNoLongerPresent"
    6007 = "NotAttemptedBecauseUserRejectedInstall"
    6008 = "NotAttemptedBecauseUserIsNotLoggedIntoAppStore"
    6009 = "NotAttemptedSupersededAppUninstallFailed"
    6010 = "NotAttemptedSupersededAppUninstallPendingReboot"
    6011 = "NotAttemptedUntargetedSupersedingAppsDetected"
    6012 = "NotAttemptedDependencyBlockedByManagedInstallerPolicy"
    6013 = "NotAttemptedUnsupportedOrIndeterminateSupersededApp"
}

# get valid Contexts(GUIDs) in the Win32Apps key
$regKeys = Get-ChildItem -Path $win32appKey -Force -ErrorAction SilentlyContinue
$getValidContexts = $regKeys | Where-Object { ([System.Guid]::TryParse((Split-Path -Path $_.Name -Leaf), [System.Management.Automation.PSReference]([guid]::empty))) } | Select-Object -ExpandProperty Name

# initialize $appKeyResults array to hold results of win32app policies found under each context (for both device and users)
$appKeyResults = @()

Foreach ($getApp in $getValidContexts) {
    $leaf = Split-Path -Path $getApp -Leaf
    $regToWork = "$win32appKey\$leaf"

    # get apps under each context
    $appKeys = Get-ChildItem -Path $regToWork -Force -ErrorAction SilentlyContinue
    $evaluatedApps = $appKeys | ForEach-Object {
        $app = $_.Name
        $appGuid = Split-Path -Path $app -Leaf
        $appReg = "$win32appKey\$leaf\$appGuid"
        $appRegValues = Get-ItemProperty -Path $appReg
        $appRegValues | Select-Object -Property @{Name = "App"; Expression = { $appGuid -replace '_\d+$' } }, @{Name = "AppReg"; Expression = { $appReg } }, @{Name = "UserObjectID"; Expression = { Split-Path -Path $regToWork -Leaf } }
    }

    # store app results in $appKeyResults array
    $appKeyResults += $evaluatedApps
}

# initialize $stateMessages array to hold results
$stateMessages = @()

Foreach ($appToWorkWith in $appKeyResults) {

    # continue if either state message key is not found
    $complianceStateKey = (Get-ItemProperty -Path "$($appToWorkWith.AppReg)\ComplianceStateMessage" -Name 'ComplianceStateMessage' -ErrorAction SilentlyContinue ).ComplianceStateMessage
    $enforcementStateKey = (Get-ItemProperty -Path "$($appToWorkWith.AppReg)\EnforcementStateMessage" -Name 'EnforcementStateMessage' -ErrorAction SilentlyContinue ).EnforcementStateMessage
   
    # if the DontTranslateStates switch is not set, attempt to translate the numerical values of the state messages to the string equivilent

    If (-not $DontTranslateStates) {
        #if the compliance state key is not found, try and replace the numerical value with the string state message equivilent

        If ($complianceStateKey) {

            # convert json object
            $complianceStateObjects = @("Applicability", "ComplianceState", "DesiredState", "TargetingMethod", "InstallContext", "TargetType")
            $complianceStateKey = $complianceStateKey | ConvertFrom-Json

            # convert values to int

            ForEach ($complianceStateObject in $complianceStateObjects) {
                $intVal = [int]$complianceStateKey.$complianceStateObject
                $tempVariable = "stateMessage$complianceStateObject"
                $stateMessageObject = (Get-Variable -Name $tempVariable).Value

                # replace values with equivalent key from state message hashtables at the top of the script

                If ($stateMessageObject[$intVal]) {
                    $complianceStateKey.$complianceStateObject = $stateMessageObject[$intVal]
                }
            }
        }
        If ($enforcementStateKey) {

            # convert json object
            $enforcementStateObjects = @("EnforcementState", "TargetingMethod")
            $enforcementStateKey = $enforcementStateKey | ConvertFrom-Json

            # convert values to int

            ForEach ($enforcementStateObject in $enforcementStateObjects) {
                $intVal2 = [int]$enforcementStateKey.$enforcementStateObject
                $tempVariable2 = "stateMessage$enforcementStateObject"
                $stateMessageObject2 = (Get-Variable -Name $tempVariable2).Value

                # replace values with equivalent key from state message hashtables at the top of the script

                If ($stateMessageObject2[$intVal2]) {
                    $enforcementStateKey.$enforcementStateObject = $stateMessageObject2[$intVal2]
                }
            }
        }
    }

    # add results to $stateMessages array
    $stateMessages += $appToWorkWith | Select-Object -Property @{Name = "UserObjectID"; Expression = { $appToWorkWith.UserObjectID } }, @{Name = "AppID"; Expression = { $appToWorkWith.App } }, @{Name = "ComplianceStateMessage"; Expression = { $complianceStateKey } }, @{Name = "EnforcementStateMessage"; Expression = { $enforcementStateKey } }, @{Name = "StateMessagesRegKey"; Expression = { $appToWorkWith.AppReg } }
}

$stateMessages