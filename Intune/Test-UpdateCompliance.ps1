$global:VortexURL = "https://v10c.events.data.microsoft.com/ping"
$global:operatingSystemName = (Get-WmiObject Win32_OperatingSystem).Name

function CheckMSAService {
    Try {
        $serviceInfo = Get-WmiObject win32_service -Filter "Name='wlidsvc'"
        $serviceStartMode = $serviceInfo.StartMode
        $serviceState = $serviceInfo.State

        if ($serviceStartMode.ToString().ToLower() -eq "disabled") {           
            Write-Output "CheckMSAService Failed: Microsoft Account Sign In Assistant Service is Disabled."
            Exit 1       
        }
        else {
            $isManualTriggeredStart = $false

            if ($serviceStartMode.ToString().TOLower() -eq "manual") {
                if (Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\wlidsvc\TriggerInfo\') {
                    $Result_CheckMSAService = "CheckMSAService: Passed. Microsoft Account Sign In Assistant Service is Manual(Triggered Start)."
                    $Exit_CheckMSAService = 0
                    $isManualTriggeredStart = $true
                }
            }

            if ($isManualTriggeredStart -eq $false) {            
                Write-Output"CheckMSAService: Failed: Microsoft Account Sign In Assistant Service is not running."
                Exit 1
            }  
        }
    }
    Catch {
        Write-Output "CheckMSAService: Failed. Exception. $($_.Exception.HResult) $($_.Exception.Message)"
        Exit 1
    }
}

function CheckUtcCsp {

    Try {
        #Check the WMI-CSP bridge (must be local system)
        $ClassName = "MDM_Win32CompatibilityAppraiser_UniversalTelemetryClient01"
        $BridgeNamespace = "root\cimv2\mdm\dmmap"
        $FieldName = "UtcConnectionReport"

        $CspInstance = get-ciminstance -Namespace $BridgeNamespace -ClassName $ClassName

        $Data = $CspInstance.$FieldName

        #Parse XML data to extract the DataUploaded field.
        $XmlData = [xml]$Data

        if (0 -eq $XmlData.ConnectionReport.ConnectionSummary.DataUploaded) {
            Write-Output "CheckUtcCsp: Failed. The only recent data uploads all failed."
            Exit 1
        }
        else {
            Write-Output "CheckUtcCsp: Passed"
        }
    }
    Catch {
        Write-Output "CheckUtcCsp: Failed. Exception. $($_.Exception.HResult) $($_.Exception.Message)"
        Exit 1
    }
}

function CheckVortexConnectivity {

    Try {   

        $Request = [System.Net.WebRequest]::Create($VortexURL)
        $Response = $Request.getResponse()

        If ($Response.StatusCode -eq 'OK') {
            Write-Output "CheckVortexConnectivity: Passed. URL $VortexURL is accessible."
        }
        Else {
            Write-Output "CheckVortexConnectivity: Failed. URL $VortexURL not accessible." 
            Exit 1

        }
    }
    Catch {
        Write-Output "CheckVortexConnectivity failed with unexpected exception. CheckVortexConnectivity. $($_.Exception.HResult) $($_.Exception.Message)"
        Exit 1
    }
}

function CheckTelemetryOptIn {

    $vCommercialIDPathPri1 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    $vCommercialIDPathPri2 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
    
    Try {
        if (($global:operatingSystemName.ToLower() -match ("^(.*?(windows )[1][0-1]*?)")) -eq $true) {
            $allowTelemetryPropertyPri1 = (Get-ItemProperty -Path $vCommercialIDPathPri1 -Name AllowTelemetry -ErrorAction SilentlyContinue).AllowTelemetry
            $allowTelemetryPropertyPri2 = (Get-ItemProperty -Path $vCommercialIDPathPri2 -Name AllowTelemetry -ErrorAction SilentlyContinue).AllowTelemetry
            
            if ($allowTelemetryPropertyPri1 -ne $null) {
                Write-Output "CheckTelemetryOptIn: Passed. AllowTelemetry property value at registry key path $vCommercialIDPathPri1 : $allowTelemetryPropertyPri1" 

                $allowTelemetryPropertyType1 = (Get-ItemProperty -Path $vCommercialIDPathPri1 -Name AllowTelemetry -ErrorAction SilentlyContinue).AllowTelemetry.gettype().Name
                if ($allowTelemetryPropertyType1 -ne "Int32") {
                    Write-Output "AllowTelemetry property value at registry key path $vCommercialIDPathPri1 is not of type REG_DWORD. It should be of type REG_DWORD."
                    Exit 1
                }

                if (-not ([int]$allowTelemetryPropertyPri1 -ge 1 -and [int]$allowTelemetryPropertyPri1 -le 3)) {
                    Write-Output "Please set the Windows telemetry level (AllowTelemetry property) to Basic (1) or above at path $vCommercialIDPathPri1. Please check https://aka.ms/uc-enrollment for more information."
                    Exit 1
                }
            }
            
            if ($allowTelemetryPropertyPri2 -ne $null) {
                Write-Output "CheckTelemetryOptIn: Passed. AllowTelemetry property value at registry key path $vCommercialIDPathPri2 : $allowTelemetryPropertyPri2"

                $allowTelemetryPropertyType2 = (Get-ItemProperty -Path $vCommercialIDPathPri2 -Name AllowTelemetry -ErrorAction SilentlyContinue).AllowTelemetry.gettype().Name
                if ($allowTelemetryPropertyType2 -ne "Int32") {
                    Write-Output "AllowTelemetry property value at registry key path $vCommercialIDPathPri2 is not of type REG_DWORD. It should be of type REG_DWORD."
                    Exit 1
                }

                if (-not ([int]$allowTelemetryPropertyPri2 -ge 1 -and [int]$allowTelemetryPropertyPri2 -le 3)) {
                    Write-Output "Please set the Windows telemetry level (AllowTelemetry property) to Basic (1) or above at path $vCommercialIDPathPri2. Check https://aka.ms/uc-enrollment for more information."
                    Exit 1
                }
            }        
        }
        else {
            Write-Output "Device must be Windows 10 or 11 to use Update Compliance."
            Exit 1
        }
    }
    Catch {
        Write-Output "CheckTelemetryOptIn failed with unexpected exception. $($_.Exception.HResult) $($_.Exception.Messag)"
        Exit 1
    }
}

function CheckCommercialId {
    Try {
        Write-Host "Start: CheckCommercialId"
        
        if (($commercialIDValue -eq $null) -or ($commercialIDValue -eq [string]::Empty)) {
            Write-Host "The commercialID parameter is incorrect. Please edit runConfig.bat and set the CommercialIDValue and rerun the script" "Error" "6" "SetupCommercialId"
            Write-Host "Script finished with error(s)" "Failure" "$global:errorCode" "ScriptEnd"
            [System.Environment]::Exit($global:errorCode)
        }

        [System.Guid]::Parse($commercialIDValue) | Out-Null

    }
    Catch {
        If (($commercialIDValue -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) -eq $false) {
            Write-Host "CommercialID mentioned in RunConfig.bat should be a GUID. It currently set to '$commercialIDValue'" "Error" "48" "CheckCommercialId"
            Write-Host "Script finished with error(s)" "Failure" "$global:errorCode" "ScriptEnd"
            [System.Environment]::Exit($global:errorCode)
        }
    }

    Write-Host "Passed: CheckCommercialId"
}

CheckMSAService
CheckUtcCsp
CheckVortexConnectivity
CheckTelemetryOptIn