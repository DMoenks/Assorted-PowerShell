# Author: Moenks, Dominik
# Version: 1.2 (02.06.2014)
# Intention: Resolves "Access Denied (0x0000052e)" errors on client computers when accessing a print server after it has been re-joined to a domain

param([Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [string[]]$ComputerNames)
$ResultCodes = @{"0" = "Successful Completion";
                "2" = "Access Denied";
                "3" = "Insufficient Privilege";
                "8" = "Unknown failure";
                "9" = "Path Not Found";
                "21" = "Invalid Parameter"}
if ($ComputerNames.Count -gt 0)
{
    $SuccessCounter = 0
    foreach ($ComputerName in $ComputerNames)
    {
        if (Test-Connection $ComputerName -Quiet)
        {
            "System $ComputerName is reachable"
            try
            {
                $LogonSession = Get-WmiObject Win32_LogonSession -ComputerName $ComputerName -ErrorAction Stop | Where-Object LogonType -EQ 2 | Sort-Object StartTime -Descending | Select-Object -First 1
                $LogonID = [Convert]::ToString($LogonSession.LogonId, 16)
                $Result = ([WMICLASS]"\\$ComputerName\ROOT\CIMV2:win32_process").Create("CMD /C NET STOP Spooler && klist -lh 0 -li $LogonID purge && NET START Spooler")
                if ($Result.ReturnValue -eq 0)
                {
                    $SuccessCounter++
                    "- Problem was resolved successfully, result was: $($ResultCodes[$Result.ReturnValue.ToString()])"
                }
                else
                {
                    "- Problem was not resolved successfully, result was: $($ResultCodes[$Result.ReturnValue.ToString()])"
                }
            }
            catch [System.UnauthorizedAccessException]
            {
                "- Problem was not resolved successfully, result was: Access Denied"
            }
        }
        else
        {
            "System $ComputerName is not reachable"
        }
    }
    "Ending the script, result was: $($ComputerNames.Count) systems checked, $SuccessCounter successful actions executed"
}
else
{
    "Ending the script, result was: 0 systems checked"
}
