$SKUtranslate = @{'STANDARDPACK' = 'Office 365 Enterprise E1';
'ENTERPRISEPACK' = 'Office 365 Enterprise E3';
'ENTERPRISEPREMIUM' = 'Office 365 Enterprise E5';
'DESKLESSPACK' = 'Office 365 Enterprise F1'}
$preferredKey = 'company'

# Input
$users = Get-AzureADUser -All $true | Where-Object{$_.AssignedLicenses.Count -gt 0}
$licenses = @{}
foreach ($user in $users)
{
    if ($null -ne ($license = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId))
    {
        if ($null -eq ($company = $user.CompanyName))
        {
            $company = ''
        }
        if ($null -eq ($location = $user.UsageLocation))
        {
            $location = ''
        }
        switch ([string]$preferredKey)
        {
            'company'
            {
                $key = $company
            }
            'location'
            {
                $key = $location
            }
            default
            {
                $key = "$location-$extension-$company"
            }
        }
        if (-not $licenses.ContainsKey($key))
        {
            $licenses.Add($key, @{})
        }
        foreach ($SKU in $license.SkuPartNumber)
        {
            if (-not $licenses[$key].ContainsKey($SKU))
            {
                $licenses[$key].Add($SKU, 0)
            }
            $licenses[$key][$SKU]++
        }
    }
}

$output = [System.Text.StringBuilder]::new()

# Output
$sortedSKUs = Get-AzureADSubscribedSku | Sort-Object -Property SkuPartNumber
$output.Append('Key') | Out-Null
foreach ($SKU in $sortedSKUs.SkuPartNumber)
{
    if ($SKUtranslate.ContainsKey($SKU))
    {
        $output.Append(";$($SKUtranslate[$SKU])") | Out-Null
    }
    else
    {
        $output.Append(';$SKU') | Out-Null
    }
}
$output.AppendLine() | Out-Null
$output.AppendLine("Overall (enabled);$($sortedSKUs.PrepaidUnits.Enabled -join ';')") | Out-Null
$output.AppendLine("Overall (consumed);$($sortedSKUs.ConsumedUnits -join ';')") | Out-Null
foreach ($key in $licenses.Keys | Sort-Object)
{
    $output.Append($key) | Out-Null
    foreach ($SKU in $sortedSKUs.SkuPartNumber)
    {
        $output.Append(";$($licenses[$key][$SKU])") | Out-Null
    }
    $output.AppendLine() | Out-Null
}
$output.ToString() | Write-Output
