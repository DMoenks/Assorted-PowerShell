# Author: Moenks, Dominik
# Version: 1.0 (06.01.2015)
# Intention: Gets the owners for all computer objects in the current domain, which normally translate to the users joining those computers to the domain

$computers =  Get-ADComputer -Filter *
$owners = @{}
foreach ($computer in $computers)
{
    $owner = (Get-Acl "AD:\$($computer.DistinguishedName)").Owner
    if ($owners.ContainsKey($owner))
    {
        $owners[$owner]++
    }
    else
    {
        $owners.Add($owner, 1)
    }
}
if (Test-Path owners.csv)
{
    Remove-Item owners.csv -Force
}
$owners.GetEnumerator() | Sort-Object -Property Value, Key -Descending | foreach {$_.Key + ";" + $_.Value >> owners.csv}
