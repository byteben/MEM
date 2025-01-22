
$Path = "C:\temp\scriptbackup"
Connect-MSGraph 
$Stream = Invoke-MSGraphRequest -Url "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts" -HttpMethod GET
 
$ScriptData = $Stream.value | Select-Object id, fileName, displayname
$TotalScripts = ($ScriptData).count
 
if ($TotalScripts -gt 0) {
    foreach ($ScriptItem in $ScriptData) {
        $Script = Invoke-MSGraphRequest -Url "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$($scriptItem.id)" -HttpMethod GET
        [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($($Script.ScriptContent))) | Out-File -FilePath $(Join-Path $Path $($Script.fileName))  -Encoding ASCII 
    }      
}