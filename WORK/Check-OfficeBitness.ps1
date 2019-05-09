$officeApps = '(excel|powerpnt|winword).exe'
$outputWidth = 20

$definition = @"
[DllImport("kernel32.dll")]
public static extern bool GetBinaryType(string lpApplicationName, ref int lpBinaryType);
"@
Add-Type -MemberDefinition $definition -name "DLL" -namespace "Import" -ErrorAction SilentlyContinue

$apps = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths')
"Checking bitness of $($apps.GetSubKeyNames() -match $officeApps -join ', ')"
foreach($app in $apps.GetSubKeyNames() -match $officeApps)
{
    $lpBinaryType = [int]::MinValue
    [Import.DLL]::GetBinaryType($apps.OpenSubKey($app).getvalue(''), [ref]$lpBinaryType) | Out-Null
    Write-Host "${app}:".PadRight($outputWidth) -NoNewline
    switch ($lpBinaryType)
    {
        0
        {
            Write-Host '32Bit'
        }
        6
        {
            Write-Host '64Bit'
        }
        default
        {
            Write-Host 'Other'
        }
    }
}
Write-Host ''
Write-Host 'Checkup finished...'
Read-Host
