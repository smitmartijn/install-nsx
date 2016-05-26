Install-NSX.ps1

This script can deploy and configure VMware NSX for you.

Usage: .\Install-NSX.ps1
  -SettingsExcel C:\install-nsx-info.xlsx                 - Mother of all information (required)
  -NSXManagerOVF C:\VMware-NSX-Manager-6.2.2-3604087.ova  - Location of the NSX Manager OVA (optional) (yes, you have to download this yourself)
  -DeployOVF             - Deploy the NSX Manager (optional)
  -RegistervCenter       - Register the NSX Manager to vCenter and SSO (optional)
  -DeployControllers     - Deploy the configured amount of NSX controllers (optional)
  -PrepareCluster        - Prepare the ESXi hosts in the vSphere cluster, configure VXLAN and add a Transport Zone (optional)
  -AddExclusions         - Add the VMs to the distributed firewall exclusion list (optional)
  -CreateLogicalSwitches - Create Logical Switches (optional)
  -CreateEdges           - Create Edge Services Gateways (optional)
  -CreateDLRs            - Create Distributed Logical Routers (optional)

Only the -SettingsExcel parameter is required and you can supply all other parameters as you see fit. To execute all
tasks, supply all paramaters. To only create Logical Switches, only use the -SettingsExcel and -CreateLogicalSwitches parameters.


Example to only deploy the NSX Manager and register it to vCenter:

PowerCLI Z:\PowerShell> .\Install-NSX.ps1 -SettingsExcel C:\install-nsx-info.xlsx -NSXManagerOVF C:\VMware-NSX-Manager-6.2.2-3604087.ova  -DeployOVF -RegistervCenter


ChangeLog:

26-05-2016 - Martijn Smit <martijn@lostdomain.org>
- Initial script
