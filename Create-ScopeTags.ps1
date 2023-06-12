<#

Requires MSAL.PS module

API permisions required:-

    DeviceManagementConfiguration.ReadWrite.All
    DeviceManagementRBAC.ReadWrite.All

#>

$authParams = @{
    ClientId    = "" #App Registration Client ID here
    TenantId    = "" #Tenant here
    Interactive = $true
}

$authToken = Get-MsalToken @authParams

$headers = @{
    "Content-Type"  = "application/json"
    "Authorization" = "$($authToken.accessToken)"
}

$scopeTags = Import-CSV .\scopeTags.csv

Foreach ($scopeTag in $scopeTags) {

    $tag = $scopeTag.Name
    $payload = @{
        "displayName" = $tag
    }

    $body = $payload | ConvertTo-Json

    $uri = "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags"
    Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body
}