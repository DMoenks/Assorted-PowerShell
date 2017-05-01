<#
.SYNOPSIS
This script is intended to migrate print queues from one or more print server to a new print server.
.DESCRIPTION
This script tries to migrate print queues from one or more print server to a new print server.
The needed print drivers need to be installed on the target server before migration. All missing drivers will be reported by the script.
.PARAMETER Sources
Specify the source server(s) from where to migrate print queues.
.PARAMETER Target
Specify the target server where to migrate print queues.
.EXAMPLE
Migrate-PrintQueues_Server.ps1 -Source "Server 1", "Server 2" -Target "Server 3"
.NOTES
Version:    1.1
Author:     Mönks, Dominik
#>

param([Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Sources,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Target)

$outputWidth = 30

# Prepare status report.
$status = @{"CHCK" = 0;
            "INST" = 0;
            "SKIPINST" = 0;
            "SKIPFAIL" = 0;}
# Check print drivers and queues already installed on the target system.
$targetDrivers = Get-PrinterDriver -ComputerName $Target | where {$_.Manufacturer -ne "Microsoft"} | select Name -Unique | foreach {$_.Name}
$targetPrinters = Get-Printer -ComputerName $Target | select Name -Unique | Out-String
$missingDrivers = @()
foreach ($source in $Sources)
{
    # Retrieve the print queues from the current source server.
    $printers = Get-Printer -ComputerName $source | where {$_.DeviceType -eq "Print" -and $_.DriverName -ne ""}
    foreach ($printer in $printers)
    {
        # Check if target server already holds a queue of the same name, then skip install.
        if ($targetPrinters.Contains($printer.Name))
        {
            $status["SKIPINST"]++
        }
        else
        {
            # Check if the target server already has a matching print installed, then migrate the print queue...
            if (($targetDrivers -like "*$($printer.DriverName)*").Count -gt 0)
            {
                Add-PrinterPort $printer.Name -PrinterHostAddress $printer.Name -ComputerName $Target -ErrorAction SilentlyContinue
                Add-Printer $printer.Name -DriverName ($targetDrivers -like "*$($printer.DriverName)*")[0] -PortName $printer.Name -Published:$true -ShareName $printer.Name -Shared:$true -ComputerName $Target
                $status["INST"]++
            }
            # ...or skip it if the driver is missing.
            else
            {
                $missingDrivers += $printer.DriverName
                $status["SKIPFAIL"]++
            }
        }
        $status["CHCK"]++
    }
}
# Report the results.
Write-Host "Checked:".PadRight($outputWidth) -NoNewline
Write-Host $status["CHCK"]
Write-Host "Installed:".PadRight($outputWidth) -NoNewline
Write-Host $status["INST"] -ForegroundColor Green
Write-Host "Skipped, already installed:".PadRight($outputWidth) -NoNewline
Write-Host $status["SKIPINST"] -ForegroundColor Yellow
Write-Host "Skipped, driver missing:".PadRight($outputWidth) -NoNewline
Write-Host $status["SKIPFAIL"] -ForegroundColor Red
Write-Host "Missing drivers:"
$missingDrivers | sort -Unique