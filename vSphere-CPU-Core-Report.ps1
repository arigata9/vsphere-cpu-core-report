# =============================================================================================================
# Script:    vSphere-CPU-Core-Report.ps1
# Version:   1.2
# Date:      02.09.2025
# By:        Markus Neuloh
# =============================================================================================================

$workdir = (Get-Location).Path

Import-Module ImportExcel

# Storing vCenters in list
$vcenters = @('vcenter1','vcenter2')

# Logging directory
$logdir = "$workdir\Logs"
$date = Get-Date -Format "yyyyMMdd"
$logfile = "$logdir\cpucorereport-$date.log"

# XLSX Output
$xlsxOutput = "$workdir\Reports\vSphere-CPU-Core-Report-$date.xlsx"

# Start transcribing for logs (stopping possible running Transcript from other processes)
Stop-Transcript | Out-Null
Start-Transcript -Path $logfile -Force -IncludeInvocationHeader

$mergedExport = @()

foreach($WLD in $vcenters)
{
    # iterate through vcenters
    Write-Host "###########################################################" -ForegroundColor Cyan
    Write-Host "Connecting to vCenter $WLD " -ForegroundColor Cyan
    Write-Host "###########################################################" -ForegroundColor Cyan

    # Connect to the vCenter server
    # To save encrypted credentials, use Get-Credential | Export-Clixml -Path "C:\temp\Cred.xml" with the windows account running this script (account will be used by DPAPI for encrypting the password)
    Connect-VIServer -Server $WLD -Credential (Import-Clixml -Path "$WLD.xml") | Out-Null

    Write-Host "Collecting data for all clusters in $WLD"

    # Get information about hosts
    $hosts = Get-VMHost | Where-Object {(Get-AdvancedSetting -Entity $_ -Name Misc.vsanWitnessVirtualAppliance).Value -eq 0} | Select-Object @{Name='Cluster';Expression={(Get-Cluster -VMHost $_).Name}},
                                    Name,
                                    @{Name='CPU Cores Sum';Expression={$_.NumCpu}},
                                    @{Name='CPU Sockets';Expression={$_.ExtensionData.Hardware.CpuInfo.NumCpuPackages}},
                                    @{Name='CPU Cores per Socket';Expression={$_.ExtensionData.Hardware.CpuInfo.NumCpuCores / $_.ExtensionData.Hardware.CpuInfo.NumCpuPackages}}


    # Add the information to the export object
    $mergedExport += @($hosts)

    # Disconnect from the vCenter server
    Disconnect-VIServer -Server $WLD -Confirm:$false

    # Clear up $hosts variable
    $hosts = $null

}

# Calculate totals
$totalCpuCores = ($mergedExport | Measure-Object -Property 'CPU Cores Sum' -Sum).Sum
$totalCpuSockets = ($mergedExport | Measure-Object -Property 'CPU Sockets' -Sum).Sum
$totalCpuCoresperSocket = ($mergedExport | Measure-Object -Property 'CPU Cores per Socket' -Sum).Sum


# Add totals row
$totals = [PSCustomObject]@{
    Cluster = 'Total'
    Name = '---'
    'CPU Cores Sum' = $totalCpuCores
    'CPU Sockets' = $totalCpuSockets
    'CPU Cores per Socket' = $totalCpuCoresperSocket
}

# Combine hosts and totals
$hostsWithTotals = $mergedExport + $totals


# Display the information in a formatted table
$hostsWithTotals | Format-Table -AutoSize

# Export data to Excel file with a new work sheet
$hostsWithTotals | Export-Excel -Path $xlsxOutput


####################################
# Send CPU core report via mail
####################################

$smtp = "smtprelay" # SMTP relay server for sending the report via mail
$from = "" # Sender email address for the report
$to = "" # Recipient email address for the report
$Cc = "" # CC email address for the report
$Bcc = "" # BCC email address for the report
$subject = "vSphere ESXi CPU Cores Monthly Report"
$body = "Report of all CPU cores from all vCenters."

Write-Host "Sending report via mail..."

Send-MailMessage -Attachments $xlsxOutput -SmtpServer $smtp -To $to -Cc $Cc -Bcc $Bcc -From $from -Subject $subject -Body $body -WarningAction Ignore


# Stop the transcript for this process
Stop-Transcript
