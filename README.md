# vsphere-cpu-core-report
Powershell script to create an Excel report for the amount of cpu cores of esxi hosts in all vCenters.
This can be run as a monthly report for example. In my case, this was required to keep track of utilized cores for VCF and Windows Datacenter licenses.

This script requires that you create encrypted vcenter credentials as per line comment #37.

To create an encrypted credential xml file, run `Get-Credential | Export-Clixml -Path "C:\temp\Cred.xml"`.

The file name should be the same as the vcenter IP/FQDN used in the `$vcenters` variable.

## Requirements

Requires the following powershell modules
- VCF.PowerCLI
- ImportExcel
