# Author: Moenks, Dominik
# Version: 1.1 (08.05.2015)
# Intention: Resolves problems with blocked, undeletable print jobs on Windows print servers

[psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::Add("ServiceStatus", "System.ServiceProcess.ServiceControllerStatus")
[psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::Add("PrinterStatus", "Microsoft.PowerShell.Cmdletization.GeneratedTypes.Printer.PrinterStatus")
[psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::Add("PrinterType", "Microsoft.PowerShell.Cmdletization.GeneratedTypes.Printer.TypeEnum")
[psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::Add("JobStatus", "Microsoft.PowerShell.Cmdletization.GeneratedTypes.PrintJob.JobStatus")

function Shutdown-Service([string]$serviceName)
{
    $self = Get-Service $serviceName
    foreach ($service in $self.DependentServices)
    {
        Shutdown-Service $service.Name
    }
    if ($self.Status -ne [ServiceStatus]::Stopped)
    {
        Write-Host "Trying to stop service $serviceName."
        Stop-Service $self
    }
}

function Startup-Service([string]$serviceName)
{
    $self = Get-Service $serviceName
    foreach ($service in $self.ServicesDependedOn)
    {
        Startup-Service $service.Name
    }
    if ($self.Status -ne [ServiceStatus]::Running)
    {
        Write-Host "Trying to start service $serviceName."
        Start-Service $self
    }
}

$jobs = @()
foreach ($printer in (Get-Printer | ?{$_.Type -eq [PrinterType]::Local -and $_.Shared -eq $true}))
{
    foreach ($printjob in (Get-PrintJob -PrinterName $printer.Name | ?{($_.JobStatus -band [JobStatus]::Error) -eq [JobStatus]::Error}))
    {
        $jobs += $printjob.ID
    }
}
Shutdown-Service Spooler
foreach ($job in $jobs)
{
    Get-ChildItem "$env:SystemRoot\system32\spool\printers\" | ?{$_.BaseName -eq $job.toString().PadLeft(5, "0")} | Remove-Item -Force
}
Startup-Service Spooler
Start-Sleep 300
Shutdown-Service Spooler
Get-ChildItem "$env:SystemRoot\system32\spool\printers\" | ?{$_.LastWriteTime -lt [DateTime]::Now.AddMinutes(-30)} | Remove-Item -Force
Startup-Service Spooler
