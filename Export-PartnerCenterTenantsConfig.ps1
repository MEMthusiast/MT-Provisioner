# Output file
$outputFile = ".\TenantsConfig.json"

# Connect to Partner Center
Connect-PartnerCenter

# Get all partner customers
$customers = Get-PartnerCustomer

Write-Host "Found $($customers.Count) tenants."

# Build tenant array
$tenantList = foreach ($customer in $customers) {
    [PSCustomObject]@{
        Name              = $customer.Name
        TenantId          = $customer.CustomerId
        UploadToAutopilot = $true
        GroupTag          = "Personal"
        OSVersion         = "Windows 11"
        OSBuild           = "25H2"
        OSEdition         = "Pro"
        OSLanguage        = "nl-nl"
        OSActivation      = "Volume"
    }
}

# Build final JSON structure
$config = @{
    Tenants = $tenantList
}

# Export JSON
$config | ConvertTo-Json -Depth 3 | Out-File $outputFile -Encoding utf8

Write-Host "Config file created: $outputFile"