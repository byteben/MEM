$Server = "SQL-SCCM"
netsh -r $Server http show sslcert IPPort=0.0.0.0:444

#$cert = (Get-ChildItem cert:\LocalMachine\My | where-object { $_.Subject -like "*$hostname*" }  | Select-Object -First 1).Thumbprint