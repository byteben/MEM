

function StartImpersonatingLoggedOnUser {

    Try {
        add-type @"
        namespace mystruct {
            using System;
            using System.Runtime.InteropServices;
                 [StructLayout(LayoutKind.Sequential)]
                 public struct WTS_SESSION_INFO
                 {
                 public Int32 SessionID;
             
                 [MarshalAs(UnmanagedType.LPStr)]
                 public String pWinStationName;
             
                 public WTS_CONNECTSTATE_CLASS State;
                 }
             
                 public enum WTS_CONNECTSTATE_CLASS
                 {
                 WTSActive,
                 WTSConnected,
                 WTSConnectQuery,
                 WTSShadow,
                 WTSDisconnected,
                 WTSIdle,
                 WTSListen,
                 WTSReset,
                 WTSDown,
                 WTSInit
                 }
                 }
            "@

        $wtsEnumerateSessions = @"
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSEnumerateSessions(
                 System.IntPtr hServer,
                 int Reserved,
                 int Version,
                 ref System.IntPtr ppSessionInfo,
                 ref int pCount);
        "@

        $wtsenum = add-type -MemberDefinition $wtsEnumerateSessions -Name PSWTSEnumerateSessions -Namespace GetLoggedOnUsers -PassThru


        $wtsqueryuserToken = @"
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSQueryUserToken(UInt32 sessionId, out System.IntPtr Token);
"@

        $wtsQuery = add-type -MemberDefinition $wtsqueryuserToken -Name PSWTSQueryServer -Namespace GetLoggedOnUsers -PassThru


        [long]$count = 0
        [long]$sessionInfo = 0
        [long]$returnValue = $wtsenum::WTSEnumerateSessions(0, 0, 1, [ref]$sessionInfo, [ref]$count)
        $datasize = [system.runtime.interopservices.marshal]::SizeOf([System.Type][mystruct.WTS_SESSION_INFO])
        $userSessionID = $null
        if ($returnValue -ne 0) {
            for ($i = 0; $i -lt $count; $i++) {
                $element = [system.runtime.interopservices.marshal]::PtrToStructure($sessionInfo + ($datasize * $i), [System.type][mystruct.WTS_SESSION_INFO])

                if ($element.State -eq [mystruct.WTS_CONNECTSTATE_CLASS]::WTSActive) {
                    $userSessionID = $element.SessionID
                }
            }

            if ($userSessionID -eq $null) {
                Write-Host "Could not impersonate Logged on user. Continuing as System. Data will be sent when a user Write-Hosts on." "Error" "41" "StartImpersonatingLoggedOnUser"
                return 
            }

            $userToken = [System.IntPtr]::Zero
            $wtsQuery::WTSQueryUserToken($userSessionID, [ref]$userToken)


            $advapiImpersonate = @'
[DllImport("advapi32.dll", SetLastError=true)]
public static extern bool ImpersonateLoggedOnUser(System.IntPtr hToken);
'@

            $impersonateUser = add-type -MemberDefinition $advapiImpersonate -Name PSImpersonategedOnUser -PassThru
            $impersonateUser::ImpersonategedOnUser($UserToken)
            $global:isImpersonatedUser = $true

            Write-Host "Passed: StartImpersonatingLoggedOnUser. Connected as Logged on user"
        }
        else {
            Write-Host "Could not impersonate Logged on user. Continuing as System. Data will be sent when a user Logs on." "Error" "41" "StartImpersonatingLoggedOnUser"
        }
    }
    Catch {
        Write-Host "StartImpersonatingLoggedOnUser failed with unexpected exception. Continuing as System. Data will be sent when a user Logs on." "Error" "42" "StartImpersonatingLoggedOnUser" $_.Exception.HResult $_.Exception.Message
    }
}

StartImpersonatingLoggedOnUser