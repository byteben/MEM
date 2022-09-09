#Simple Script to detect a Ninja Registry Value
$RegPath = "HKLM:\Software\Ninja\Offers"
$RegName = "WeAllGetFreeStuff"
$RegValue = "Yes"

Try {
    If (Test-Path $RegPath) {
        $RegResult = Get-ItemProperty $RegPath -Name $RegName -ErrorAction Stop | Select-Object -ExpandProperty $RegName

        If ($RegResult -eq $RegValue) {
            Write-Output "Match. Expected $($RegValue) and got $($RegResult)"
            Exit 0 
        }
        else {
            If ($RegResult::IsNullOrEmpty) {
                $RegResult = "<<null>>"
            }
            Write-Output "No Match. Expected $($RegValue) but got $($RegResult)"
            Exit 1
        }
    }
    else {
        Write-Output "No Match. Expected $($RegPath) not found"
        Exit 1
    }
}
Catch {
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    Exit 1
}
