#Simple Script to remediate a Ninja Registry Value
$RegPath = "HKLM:\Software\Ninja\Offers"
$RegName = "WeAllGetFreeStuff"
$RegValue = "Yes"
$RegType = "String"

Try {
    New-Item $RegPath -Force | New-ItemProperty -Name $RegName -Value $RegValue -PropertyType $RegType -Force | Out-Null
    #Exit 0
}
Catch {
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    #Exit 1 
}