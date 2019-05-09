param([Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Computername)
$files = "System32\ctfmon.exe","SysWOW64\ctfmon.exe"
$owner = [System.Security.Principal.NTAccount]::new("BUILTIN", "Administrators")

Write-Host "Checking for connection to device ${computername}:".PadRight(75) -NoNewline
if (Test-Connection $Computername -Quiet)
{
    Write-Host "Succeeded" -ForegroundColor Green
    foreach ($file in $files)
    {
        $path = "\\$Computername\C$\Windows"
        Write-Host "Checking for file $path\${file}:".PadRight(75) -NoNewline
        if (Test-Path "$path\$file")
        {
            Write-Host "Succeeded" -ForegroundColor Green
            
            Write-Host "Taking ownership of file..."
            $ACL = Get-Acl "$path\$file"
            $ACL.SetOwner($owner)
            Set-Acl "$path\$file" $ACL
            
            Write-Host "Setting ACL for file..."
            $ACL = Get-Acl "$path\$file"
            $rule = [System.Security.AccessControl.FileSystemAccessRule]::new($owner, [Security.AccessControl.FileSystemRights]::Modify, [Security.AccessControl.AccessControlType]::Allow)
            $ACL.AddAccessRule($rule)            
            Set-Acl "$path\$file" $ACL
            
            Write-Host "Replacing file..."
            New-Item "$path\$file" -ItemType File -Force | Out-Null
        }
        else
        {
            Write-Host "Failed" -ForegroundColor Yellow -NoNewline
            Write-Host ", skipping file"
        }
    }
}
else
{
    Write-Host "Failed" -ForegroundColor Red
}
