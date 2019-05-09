Connect-MsolService
$users = Get-MsolUser -All | ?{$_.IsLicensed -eq $true}
foreach ($user in $users)
{
    foreach ($license in $user.LicenseAssignmentDetails)
    {
        if (($license.Assignments | ?{$_.ReferencedObjectId -ne $user.ObjectId}).Count -gt 0 -and
            ($license.Assignments | ?{$_.ReferencedObjectId -eq $user.ObjectId}).Count -gt 0)
        {
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses "$($license.AccountSku.AccountName):$($license.AccountSku.SkuPartNumber)"
        }
    }
}
