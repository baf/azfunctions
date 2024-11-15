# credit to https://practical365.com/using-azure-functions-for-exchange-online/ for the building blocks

using namespace System.Net 

# Input bindings are passed in via param block. 
param($Request, $TriggerMetadata) 

Write-Host "Connecting to EXO..." 
$paramsEXO = @{ 
    ManagedIdentity = $true 
    Organization = 'TENANTNAME.onmicrosoft.com' # your tenant name here
    ShowBanner = $false 
    CommandName = @('Get-EXOMailbox','Set-Mailbox') 
    ErrorAction = 'Stop' 
} 
try { 
    Connect-ExchangeOnline @paramsEXO 
} 
catch{ 
    # create response body in JSON format 
    $body = $_.Exception.Message | ConvertTo-Json -Compress -Depth 10 
    break 
} 
# get variables from query parameters
$mailbox = $Request.Query.Mailbox
$address = $Request.Query.Address
$domain = $Request.Query.Domain

if (-not [System.String]::IsNullOrEmpty($mailbox) -and [System.String]::IsNullOrEmpty($address)) 
{ 
    try { 
        Write-Host "Retrieving proxyAddresses for $mailbox" 
        $proxyAddressesCount = (Get-EXOMailbox -Identity $mailbox -Properties EmailAddresses -ErrorAction Stop | Select-Object -ExpandProperty EmailAddresses | Measure-Object).Count

        if ($proxyAddressesCount -ge 300) # 300 addresses is the documented limit https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#sending-limits-1
        {
            Write-Host "proxyAddresses exceed 300 - clearing old addresses"
            # Clear all proxyAddresses matching the specified domain
            Get-EXOMailbox $mailbox -Properties EmailAddresses | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_ -like "*@$domain" } | ForEach-Object { Set-Mailbox -Identity $mailbox -EmailAddresses @{remove=$_} }
        }
        Write-Host "Adding new proxyAddress $address to mailbox $mailbox"
        # Add the new proxyAddress
        Set-Mailbox -Identity $mailbox -EmailAddresses @{add=$address} -ErrorAction Stop        
    } 
    catch{ 
        # create response body in JSON format 
        $body = $_.Exception.Message | ConvertTo-Json -Compress -Depth 10 
    } 
} 

# Associate values to output bindings by calling 'Push-OutputBinding'. 
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ 
    StatusCode = [HttpStatusCode]::OK 
    Body = $body 
})