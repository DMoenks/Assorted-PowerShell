# Author: Moenks, Dominik
# Version: 1.2 (11.02.2014)
# Intention: Give "Modify" rights to a specified domain group or user for a specified folder and all subfolders

param([string]$ADObjectName, [string]$ADDomain, [string]$FolderName)

if ($ADObjectName -ne $null -and $ADObjectName -ne "" -and $ADDomain -ne $null -and $ADDomain -ne "" -and $FolderName -ne $null -and $FolderName -ne "")
{
    if (Test-Path $FolderName)
    {
        $ADObject = $null
        try
        {
            $ADObject = Get-ADUser $ADObjectName -Server $ADDomain
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException],
        [Microsoft.ActiveDirectory.Management.ADMultipleMatchingIdentitiesException],
        [Microsoft.ActiveDirectory.Management.ADServerDownException]
        {

        }
        try
        {
            $ADObject = Get-ADGroup $ADObjectName -Server $ADDomain
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException],
        [Microsoft.ActiveDirectory.Management.ADMultipleMatchingIdentitiesException],
        [Microsoft.ActiveDirectory.Management.ADServerDownException]
        {

        }
        if ($ADObject -ne $null)
        {
            $Folders = Get-ChildItem -Path $FolderName -Recurse -Directory
            $FolderCounter = 0
            foreach ($Folder in $Folders)
            {
                $ACL = Get-Acl $Folder.FullName
        
                $ACLRule = New-Object System.Security.AccessControl.FileSystemAccessRule($ADObject.SID, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
                $ACL.AddAccessRule($ACLRule)
        
                Set-Acl $Folder.FullName $ACL

                $FolderCounter++
                "Fortschritt: " + $FolderCounter + "/" + $Folders.Count
            }
        }
        else
        {
            "First parameter value was either no distinct user/group name or the second parameter value was no domain name."
        }
    }
    else
    {
        "Third parameter value was no correct folder name or the folder doesn't exist."
    }
}
else
{
    "Please specify a user/group name as first parameter value, a domain name as second parameter value and a folder as third parameter value."
}
