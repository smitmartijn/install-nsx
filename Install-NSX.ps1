#
# Install-NSX.ps1
#
# This script can deploy and configure VMware NSX for you.
#
# Usage: .\Install-NSX.ps1
#   -SettingsExcel C:\install-nsx-info.xlsx                 - Mother of all information (required)
#   -NSXManagerOVF C:\VMware-NSX-Manager-6.2.2-3604087.ova  - Location of the NSX Manager OVA (optional) (yes, you have to download this yourself)
#   -DeployOVF             - Deploy the NSX Manager (optional)
#   -RegistervCenter       - Register the NSX Manager to vCenter and SSO (optional)
#   -InsertLicense         - Insert the NSX License into vCenter (required in 6.2.3+ before host prep)
#   -DeployControllers     - Deploy the configured amount of NSX controllers (optional)
#   -PrepareCluster        - Prepare the ESXi hosts in the vSphere cluster, configure VXLAN and add a Transport Zone (optional)
#   -AddExclusions         - Add the VMs to the distributed firewall exclusion list (optional)
#   -CreateLogicalSwitches - Create Logical Switches (optional)
#   -CreateEdges           - Create Edge Services Gateways (optional)
#   -CreateDLRs            - Create Distributed Logical Routers (optional)
#
# Only the -SettingsExcel parameter is required and you can supply all other parameters as you see fit. To execute all
# tasks, supply all paramaters. To only create Logical Switches, only use the -SettingsExcel and -CreateLogicalSwitches parameters.
#
#
# Example to only deploy the NSX Manager and register it to vCenter:
#
# PowerCLI Z:\PowerShell> .\Install-NSX.ps1 -SettingsExcel C:\install-nsx-info.xlsx -NSXManagerOVF C:\VMware-NSX-Manager-6.2.2-3604087.ova  -DeployOVF -RegistervCenter
#
#
# ChangeLog:
#
# 26-05-2016 - Martijn Smit <martijn@lostdomain.org>
# - Initial script
#
#
param (
  [parameter(Mandatory=$true, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$SettingsExcel = "",
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$NSXManagerOVF = "",
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$DeployOVF,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$RegistervCenter,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$InsertLicense,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$DeployControllers,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$PrepareCluster,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$AddExclusions,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$CreateLogicalSwitches,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$CreateEdges,
  [parameter(Mandatory=$false, ValueFromPipeLine=$true, ValueFromPipeLineByPropertyName=$true)]
  [ValidateNotNullOrEmpty()]
  [switch]$CreateDLRs
)

$stopwatch = [system.diagnostics.stopwatch]::startNew()

# Load PowerCLI
if (!(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
  if (Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI' ) {
    $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI'
  }
  else {
    $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\VMware, Inc.\VMware vSphere PowerCLI'
  }
  .(join-path -path (Get-ItemProperty  $Regkey).InstallPath -childpath 'Scripts\Initialize-PowerCLIEnvironment.ps1')
}
if (!(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
  Write-Host "VMware modules not loaded/unable to load"
  Exit
}

# Load PowerNSX
Import-Module -Name '.\PowerNSX.psm1' -ErrorAction SilentlyContinue -DisableNameChecking

# Load the distributed switch module

# We need to point the PSModulePath variable to the PowerCLI modules directory,
# for some reason the PowerCLI installer doesn't do this for us:
$p = [Environment]::GetEnvironmentVariable("PSModulePath");
$p += ";C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Modules\";
[Environment]::SetEnvironmentVariable("PSModulePath",$p);
# Now import distributed vSwitch module!
Import-Module VMware.VimAutomation.Vds -ErrorAction Stop

# Import some supporting functions
.".\Install-NSX-Functions.ps1"

# Check if the excel exists
if(!(Test-Path $SettingsExcel)) {
  Write-Host "Settings Excel file '$SettingsExcel' not found!" -ForegroundColor "red"
  Exit
}


$Excel = New-Object -COM "Excel.Application"
$Excel.Visible = $False
$WorkBook = $Excel.Workbooks.Open($SettingsExcel)
$WorkSheet = $WorkBook.Sheets.Item(1)


$NSX_MGR_Name      = $WorkSheet.Cells.Item(3, 1).Value()
$NSX_MGR_Hostname  = $WorkSheet.Cells.Item(3, 2).Value()
$NSX_MGR_IP        = $WorkSheet.Cells.Item(3, 3).Value()
$NSX_MGR_Netmask   = $WorkSheet.Cells.Item(3, 4).Value()
$NSX_MGR_Gateway   = $WorkSheet.Cells.Item(3, 5).Value()
$NSX_MGR_DNSServer = $WorkSheet.Cells.Item(3, 6).Value()
$NSX_MGR_DNSDomain = $WorkSheet.Cells.Item(3, 7).Value()
$NSX_MGR_NTPServer = $WorkSheet.Cells.Item(3, 8).Value()
$NSX_MGR_CLI_Pass  = $WorkSheet.Cells.Item(3, 9).Value()


$NSX_VC_IP        = $WorkSheet.Cells.Item(7, 1).Value()
$NSX_VC_Username  = $WorkSheet.Cells.Item(7, 2).Value()
$NSX_VC_Password  = $WorkSheet.Cells.Item(7, 3).Value()
$NSX_VC_Cluster   = $WorkSheet.Cells.Item(7, 4).Value()
$NSX_VC_Network   = $WorkSheet.Cells.Item(7, 5).Value()
$NSX_VC_Datastore = $WorkSheet.Cells.Item(7, 6).Value()
$NSX_VC_Folder    = $WorkSheet.Cells.Item(7, 7).Value()
$NSX_License      = $WorkSheet.Cells.Item(7, 8).Value()

$NSX_VC_Connect_IP        = $WorkSheet.Cells.Item(11, 1).Value()
$NSX_VC_Connect_Username  = $WorkSheet.Cells.Item(11, 2).Value()
$NSX_VC_Connect_Password  = $WorkSheet.Cells.Item(11, 3).Value()

$NSX_Controllers_Cluster = $WorkSheet.Cells.Item(15, 1).Value()
$NSX_Controllers_Datastore = $WorkSheet.Cells.Item(15, 2).Value()
$NSX_Controllers_PortGroup = $WorkSheet.Cells.Item(15, 3).Value()
$NSX_Controllers_Password = $WorkSheet.Cells.Item(15, 4).Value()
$NSX_Controllers_Amount = [int]$WorkSheet.Cells.Item(15, 5).Value()

$NSX_VXLAN_Cluster = $WorkSheet.Cells.Item(19, 1).Value()
$NSX_VXLAN_DSwitch = $WorkSheet.Cells.Item(19, 2).Value()
$NSX_VXLAN_VLANID = $WorkSheet.Cells.Item(19, 3).Value()
$NSX_VXLAN_VTEP_Count = $WorkSheet.Cells.Item(19, 4).Value()

$NSX_VXLAN_Segment_ID_Begin = [int]$WorkSheet.Cells.Item(19, 5).Value()
$NSX_VXLAN_Segment_ID_End = [int]$WorkSheet.Cells.Item(19, 6).Value()

$NSX_VXLAN_Multicast_Range_Begin = $WorkSheet.Cells.Item(19, 7).Value()
$NSX_VXLAN_Multicast_Range_End = $WorkSheet.Cells.Item(19, 8).Value()

$NSX_VXLAN_Failover_Mode = $WorkSheet.Cells.Item(21, 1).Value()
$NSX_VXLAN_MTU_Size = $WorkSheet.Cells.Item(21, 2).Value()

$NSX_VXLAN_TZ_Name = $WorkSheet.Cells.Item(25, 1).Value()
$NSX_VXLAN_TZ_Mode = $WorkSheet.Cells.Item(25, 2).Value()

$a = Release-Ref($WorkSheet)
$WorkSheet_IPPools = $WorkBook.Sheets.Item(2)

$NSX_Controllers_IP_Pool_Name      = $WorkSheet_IPPools.Cells.Item(2, 2).Value()
$NSX_Controllers_IP_Pool_Gateway   = $WorkSheet_IPPools.Cells.Item(2, 3).Value()
$NSX_Controllers_IP_Pool_Prefix    = $WorkSheet_IPPools.Cells.Item(2, 4).Value()
$NSX_Controllers_IP_Pool_DNS1      = $WorkSheet_IPPools.Cells.Item(2, 5).Value()
$NSX_Controllers_IP_Pool_DNS2      = $WorkSheet_IPPools.Cells.Item(2, 6).Value()
$NSX_Controllers_IP_Pool_DNSSuffix = $WorkSheet_IPPools.Cells.Item(2, 7).Value()
$NSX_Controllers_IP_Pool_Start     = $WorkSheet_IPPools.Cells.Item(2, 8).Value()
$NSX_Controllers_IP_Pool_End       = $WorkSheet_IPPools.Cells.Item(2, 9).Value()

$NSX_VXLAN_IP_Pool_Name      = $WorkSheet_IPPools.Cells.Item(3, 2).Value()
$NSX_VXLAN_IP_Pool_Gateway   = $WorkSheet_IPPools.Cells.Item(3, 3).Value()
$NSX_VXLAN_IP_Pool_Prefix    = $WorkSheet_IPPools.Cells.Item(3, 4).Value()
$NSX_VXLAN_IP_Pool_DNS1      = $WorkSheet_IPPools.Cells.Item(3, 5).Value()
$NSX_VXLAN_IP_Pool_DNS2      = $WorkSheet_IPPools.Cells.Item(3, 6).Value()
$NSX_VXLAN_IP_Pool_DNSSuffix = $WorkSheet_IPPools.Cells.Item(3, 7).Value()
$NSX_VXLAN_IP_Pool_Start     = $WorkSheet_IPPools.Cells.Item(3, 8).Value()
$NSX_VXLAN_IP_Pool_End       = $WorkSheet_IPPools.Cells.Item(3, 9).Value()

$a = Release-Ref($WorkSheet_IPPools)




if($DeployOVF.IsPresent)
{
  Write-Host "Starting NSX Manager deployment..." -ForegroundColor "yellow"
  # Test if OVA exists
  if (Test-Path $NSXManagerOVF)
  {
    # Connect to vCenter first
    # Ignore SSL warnings
    $ignore = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False -ErrorAction SilentlyContinue -Scope User
    if(!(Connect-VIServer -Server $NSX_VC_IP -User $NSX_VC_Username -Password $NSX_VC_Password)) {
      Write-Host "Unable to connect to vCenter!" -ForegroundColor "red"
      Exit
    }

    # Then use PowerNSX to deploy OVA
    if(!(New-NSXManager -NsxManagerOVF $NSXManagerOVF -Name $NSX_MGR_Name -ClusterName $NSX_VC_Cluster -ManagementPortGroupName $NSX_VC_Network -DatastoreName $NSX_VC_Datastore -FolderName $NSX_VC_Folder -CliPassword $NSX_MGR_CLI_Pass -CliEnablePassword $NSX_MGR_CLI_Pass -Hostname $NSX_MGR_Hostname -IpAddress $NSX_MGR_IP -Netmask $NSX_MGR_Netmask -Gateway $NSX_MGR_Gateway -DnsServer $NSX_MGR_DNSServer -DnsDomain $NSX_MGR_DNSDomain -NtpServer $NSX_MGR_NTPServer -EnableSsh -StartVm)) {
      Write-Host "Unable to deploy NSX Manager OVF! $_" -ForegroundColor "red"
      Exit
    }

    Write-Host "Deployed NSX Manager!" -ForegroundColor "green"
    $disconnect = Disconnect-VIServer -Confirm:$False

    Write-Host -NoNewLine "If you have other tasks for me, wait for the NSX Management Service to come online and press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
  }
  else
  {
    # If the OVA doesn't exist, fail
    Write-Host "The file $NSXManagerOVF doesn't exist!" -ForegroundColor "red"
    Exit
  }
}

if($RegistervCenter.IsPresent)
{
  Write-Host "Connecting to NSX Manager and registering it to vCenter..." -ForegroundColor "yellow"

  # Connect to vCenter first
  if(!(Connect-VIServer -Server $NSX_VC_Connect_IP -User $NSX_VC_Connect_Username -Password $NSX_VC_Connect_Password)) {
    Write-Host "Unable to connect to vCenter!" -ForegroundColor "red"
    Exit
  }

  # Connect to NSX Manager
  if(!(Connect-NSXServer -Server $NSX_MGR_IP -Username admin -Password $NSX_MGR_CLI_Pass)) {
    Write-Host "Unable to connect to NSX Manager!" -ForegroundColor "red"
    Exit
  }
  # Configure vCenter connection on NSX Manager
  $vc = Set-NSXManager -vCenterServer $NSX_VC_Connect_IP -vCenterUserName $NSX_VC_Connect_Username -vCenterPassword $NSX_VC_Connect_Password
  # Configure vCenter SSO on NSX Manager
  $sso = Set-NSXManager -SsoServer $NSX_VC_Connect_IP -SsoUserName $NSX_VC_Connect_Username -SsoPassword $NSX_VC_Connect_Password -AcceptAnyThumbprint
  Write-Host "Configured vCenter and SSO on NSX Manager!" -ForegroundColor "green"
}

# Thanks to Anthony Burke for the upcoming license bit!
# t. @pandom / w. http://networkinferno.net/license-nsx-via-automation-with-powercli
if($InsertLicense.IsPresent)
{
  Write-Host "Assigning NSX license to vCenter..." -ForegroundColor "yellow"
  $ServiceInstance = Get-View ServiceInstance
  $LicenseManager = Get-View $ServiceInstance.Content.licenseManager
  $LicenseAssignmentManager = Get-View $LicenseManager.licenseAssignmentManager
  $output = $LicenseAssignmentManager.UpdateAssignedLicense("nsx-netsec", $NSX_License, $NULL)

  # Check if license has been properly set
  $CheckLicense = $LicenseAssignmentManager.QueryAssignedLicenses("nsx-netsec")
  if($CheckLicense.AssignedLicense.LicenseKey -ne $NSX_License) {
    Write-Host "Setting the NSX License failed! Error: $CheckLicense" -ForegroundColor "red"
    Exit
  }
  else {
    Write-Host "Configured NSX License on vCenter!" -ForegroundColor "green"
  }
}

if($DeployControllers.IsPresent)
{
  Write-Host "Deploying NSX Controllers ($NSX_Controllers_Amount)...(this will take a while)" -ForegroundColor "yellow"

  # NSX Controllers sanity check
  if($NSX_Controllers_Amount -gt 3) {
    Write-Host "NSX Controllers cannot be more than 3!" -ForegroundColor "red"
    Exit
  }

  # Check if the IP Pool already exists
  $ippool = Get-NsxIpPool -Name $NSX_Controllers_IP_Pool_Name -ErrorAction SilentlyContinue
  if($ippool -eq $null) {
    # if the IP Pool doesn't exist, create it!
    $ippool = New-NsxIpPool -Name $NSX_Controllers_IP_Pool_Name -Gateway $NSX_Controllers_IP_Pool_Gateway -SubnetPrefixLength $NSX_Controllers_IP_Pool_Prefix -DnsServer1 $NSX_Controllers_IP_Pool_DNS1 -DnsServer2 $NSX_Controllers_IP_Pool_DNS2 -DnsSuffix $NSX_Controllers_IP_Pool_DNSSuffix -StartAddress $NSX_Controllers_IP_Pool_Start -EndAddress $NSX_Controllers_IP_Pool_End
  }

  # Create controllers
  $cluster = Get-Cluster -Name $NSX_Controllers_Cluster
  $datastore = Get-Datastore -Name $NSX_Controllers_Datastore
  $portgroup = Get-VirtualPortGroup -Name $NSX_Controllers_PortGroup -Distributed
  $i = 1
  While ($i -le $NSX_Controllers_Amount) {
    Write-Host "Deloying controller $i ..." -ForegroundColor "yellow"
    $controller = New-NsxController -Confirm:$False -IpPool $ippool -Cluster $cluster -Datastore $datastore -PortGroup $portgroup -Password $NSX_Controllers_Password -Wait
    $i += 1
  }

  Write-Host "Controllers deployed!" -ForegroundColor "green"
}

if($PrepareCluster.IsPresent)
{
  Write-Host "Preparing vSphere cluster $NSX_VXLAN_Cluster (installing VIBs)..." -ForegroundColor "yellow"
  $cluster = Get-Cluster -Name $NSX_VXLAN_Cluster
  $prep = Install-NsxCluster -Cluster $cluster -VxlanPrepTimeout 300
  Write-Host "Configuring VXLAN on cluster $NSX_VXLAN_Cluster " -ForegroundColor "yellow"

  # Configure VXLAN Segment ID
  Write-Host "Adding a segment ID range.." -ForegroundColor "yellow"
  $ignorethis = New-NsxSegmentIdRange -Name "Segment1" -Begin $NSX_VXLAN_Segment_ID_Begin -End $NSX_VXLAN_Segment_ID_End

  # Configure VXLAN Multicast IP range
  Write-Host "Adding a multicast IP range.." -ForegroundColor "yellow"
  $ignorethis = New-NsxMulticastRange -Name "Multicast1" -Begin $NSX_VXLAN_Multicast_Range_Begin -End $NSX_VXLAN_Multicast_Range_End

  $vds = Get-VDSwitch -Name $NSX_VXLAN_DSwitch
  $ignorethis = New-NsxVdsContext -VirtualDistributedSwitch $vds -Teaming $NSX_VXLAN_Failover_Mode -Mtu $NSX_VXLAN_MTU_Size

  Write-Host "Adding the VXLAN IP Pool.." -ForegroundColor "yellow"
  # Check if the IP Pool already exists
  $ippool = Get-NsxIpPool -Name $NSX_VXLAN_IP_Pool_Name -ErrorAction SilentlyContinue
  if($ippool -eq $null) {
    # if the IP Pool doesn't exist, create it!
    $ippool = New-NsxIpPool -Name $NSX_VXLAN_IP_Pool_Name -Gateway $NSX_VXLAN_IP_Pool_Gateway -SubnetPrefixLength $NSX_VXLAN_IP_Pool_Prefix -DnsServer1 $NSX_VXLAN_IP_Pool_DNS1 -DnsServer2 $NSX_VXLAN_IP_Pool_DNS2 -DnsSuffix $NSX_VXLAN_IP_Pool_DNSSuffix -StartAddress $NSX_VXLAN_IP_Pool_Start -EndAddress $NSX_VXLAN_IP_Pool_End
  }

  Write-Host "Configuring VXLAN VTEPs on the cluster.." -ForegroundColor "yellow"

  $ignorethis = New-NsxClusterVxlanConfig -Cluster $cluster -VirtualDistributedSwitch $vds -IpPool $ippool -VlanId $NSX_VXLAN_VLANID -VtepCount $NSX_VXLAN_VTEP_Count

  # Add a transport zone
  Write-Host "Adding transport zone.." -ForegroundColor "yellow"
  $ignorethis = New-NsxTransportZone -Cluster $cluster -Name $NSX_VXLAN_TZ_Name -ControlPlaneMode $NSX_VXLAN_TZ_Mode

  Write-Host "Cluster prepared!" -ForegroundColor "green"
}

if($AddExclusions.IsPresent)
{
  Write-Host "Adding VM exclusions from the distributed firewall.." -ForegroundColor "yellow"
  # The exclusions are on sheet 3 of the excel
  $WorkSheet_Exclusions = $WorkBook.Sheets.Item(3)
  # Start at row 3 (minus headers) and loop through them while the cells are not empty
  $intRow = 3
  $created = 0
  While ($WorkSheet_Exclusions.Cells.Item($intRow, 1).Value() -ne $null)
  {
    # Get the VM name and add the VM to the exclusion list
    $VM_Name = $WorkSheet_Exclusions.Cells.Item($intRow, 1).Value()
    if(!(Add-NsxFirewallExclusionList -VirtualMachine (Get-VM -Name $VM_Name))) {
      Write-Host "Unable to add exclusion for $VM_Name $_" -ForegroundColor "red"
    }
    else {
      $created++
    }
    $intRow++
  }
  $a = Release-Ref($WorkSheet_Exclusions)
  Write-Host "Added $created VMs to exclusion list!" -ForegroundColor "green"
}

if($CreateLogicalSwitches.IsPresent)
{
  Write-Host "Creating Logical Switches.." -ForegroundColor "yellow"
  # We need the vdnscope ID of the transport zone for this!
  $scopeId = Get-NsxTransportZone -Name $NSX_VXLAN_TZ_Name

  if($scopeId -eq $null) {
    Write-Host "NSX Transport Zone not found, have you prepared the cluster?" -ForegroundColor "red"
    Exit
  }

  # The logical switches are on sheet 4 of the excel
  $WorkSheet_LS = $WorkBook.Sheets.Item(4)
  # Start at row 2 (minus header) and loop through them while the cells are not empty
  $intRow = 2
  $created = 0
  While ($WorkSheet_LS.Cells.Item($intRow, 1).Value() -ne $null)
  {
    # Get the LS name and add it to NSX
    $LS_Name = $WorkSheet_LS.Cells.Item($intRow, 1).Value()
    $LS_Desc = $WorkSheet_LS.Cells.Item($intRow, 2).Value()

    if(!(New-NsxLogicalSwitch -Name $LS_Name -Description $LS_Desc -vdnScope $scopeId)) {
      Write-Host "Unable to create Logical Switch $LS_Name $_" -ForegroundColor "red"
    }
    else {
      $created++
    }
    $intRow++
  }
  $a = Release-Ref($WorkSheet_LS)
  Write-Host "Added $created Logical Switches!" -ForegroundColor "green"
}

if($CreateEdges.IsPresent)
{
  Write-Host "Creating Edge Services Gateways.." -ForegroundColor "yellow"

  # The ESGs are on sheet 5 of the excel
  $WorkSheet_ESG = $WorkBook.Sheets.Item(5)
  # Start at row 2 (minus header) and loop through them while the cells are not empty
  $intRow = 2
  $created = 0
  While ($WorkSheet_ESG.Cells.Item($intRow, 1).Value() -ne $null)
  {
    $Edge_Name       = $WorkSheet_ESG.Cells.Item($intRow, 1).Value()
    $Edge_Cluster    = $WorkSheet_ESG.Cells.Item($intRow, 2).Value()
    $Edge_Datastore  = $WorkSheet_ESG.Cells.Item($intRow, 3).Value()
    $Edge_Password   = $WorkSheet_ESG.Cells.Item($intRow, 4).Value()
    $Edge_FormFactor = $WorkSheet_ESG.Cells.Item($intRow, 5).Value()
    $Edge_HA         = $WorkSheet_ESG.Cells.Item($intRow, 6).Value()
    $Edge_Hostname   = $WorkSheet_ESG.Cells.Item($intRow, 7).Value()

    $Edge_VNIC0_Name         = $WorkSheet_ESG.Cells.Item($intRow, 8).Value()
    $Edge_VNIC0_IP           = $WorkSheet_ESG.Cells.Item($intRow, 9).Value()
    $Edge_VNIC0_Prefixlength = $WorkSheet_ESG.Cells.Item($intRow, 10).Value()
    $Edge_VNIC0_PortGroup    = $WorkSheet_ESG.Cells.Item($intRow, 11).Value()

    $enableHA = $false
    if($Edge_HA -eq "Yes") {
      $enableHA = $true
    }

    # figure out the connected portgroup. First, assume it's a logical switch and if it's not, move on to a PortGroup
    $connectedTo = (Get-NsxTransportZone -Name $NSX_VXLAN_TZ_Name | Get-NsxLogicalSwitch $Edge_VNIC0_PortGroup)
    if($connectedTo -eq $null) {
      $connectedTo = (Get-VDPortgroup $Edge_VNIC0_PortGroup)
    }

    $vnic0 = New-NsxEdgeInterfaceSpec -Index 0 -Name $Edge_VNIC0_Name -Type Uplink -ConnectedTo $connectedTo -PrimaryAddress $Edge_VNIC0_IP -SubnetPrefixLength $Edge_VNIC0_Prefixlength
    $edge = New-NsxEdge -Name $Edge_Name -Cluster (Get-Cluster -Name $Edge_Cluster) -Datastore (Get-Datastore -Name $Edge_Datastore) -FormFactor $Edge_FormFactor -Password $Edge_Password -Hostname $Edge_Hostname -EnableHa:$enableHA -Interface $vnic0

    if(!($edge)) {
      Write-Host "Unable to create Edge $Edge_Name $_" -ForegroundColor "red"
    }
    else {
      $created++
    }
    $intRow++
  }
  $a = Release-Ref($WorkSheet_ESG)
  Write-Host "Added $created Edge Services Gateways!" -ForegroundColor "green"
}








if($CreateDLRs.IsPresent)
{
  Write-Host "Creating Distributed Logical Routers.." -ForegroundColor "yellow"

  # The distributed routers are on sheet 6 of the excel
  $WorkSheet_DLR = $WorkBook.Sheets.Item(6)
  # Start at row 2 (minus header) and loop through them while the cells are not empty
  $intRow = 2
  $created = 0
  While ($WorkSheet_DLR.Cells.Item($intRow, 1).Value() -ne $null)
  {
    $DLR_Name       = $WorkSheet_DLR.Cells.Item($intRow, 1).Value()
    $DLR_Cluster    = $WorkSheet_DLR.Cells.Item($intRow, 2).Value()
    $DLR_Datastore  = $WorkSheet_DLR.Cells.Item($intRow, 3).Value()
    $DLR_Password   = $WorkSheet_DLR.Cells.Item($intRow, 4).Value()
    $DLR_HA         = $WorkSheet_DLR.Cells.Item($intRow, 5).Value()

    $DLR_VNIC0_Name         = $WorkSheet_DLR.Cells.Item($intRow, 6).Value()
    $DLR_VNIC0_IP           = $WorkSheet_DLR.Cells.Item($intRow, 7).Value()
    $DLR_VNIC0_Prefixlength = $WorkSheet_DLR.Cells.Item($intRow, 8).Value()
    $DLR_VNIC0_PortGroup    = $WorkSheet_DLR.Cells.Item($intRow, 9).Value()

    $DLR_MGMT_PortGroup     = $WorkSheet_DLR.Cells.Item($intRow, 10).Value()

    $enableHA = $false
    if($Edge_HA -eq "Yes") {
      $enableHA = $true
    }

    # figure out the connected portgroup. First, assume it's a logical switch and if it's not, move on to a PortGroup
    $connectedTo = (Get-NsxTransportZone -Name $NSX_VXLAN_TZ_Name | Get-NsxLogicalSwitch $DLR_VNIC0_PortGroup)
    if($connectedTo -eq $null) {
      $connectedTo = (Get-VDPortgroup $DLR_VNIC0_PortGroup)
    }
    # same for the management port
    $mgtNic = (Get-NsxTransportZone -Name $NSX_VXLAN_TZ_Name | Get-NsxLogicalSwitch $DLR_MGMT_PortGroup)
    if($mgtNic -eq $null) {
      $mgtNic = (Get-VDPortgroup $DLR_MGMT_PortGroup)
    }

    $vnic0 = New-NsxLogicalRouterInterfaceSpec -Name $DLR_VNIC0_Name -Type Uplink -ConnectedTo $connectedTo -PrimaryAddress $DLR_VNIC0_IP -SubnetPrefixLength $DLR_VNIC0_Prefixlength
    $dlr = New-NsxLogicalRouter -Name $DLR_Name -Cluster (Get-Cluster -Name $DLR_Cluster) -Datastore (Get-Datastore -Name $DLR_Datastore) -EnableHa:$enableHA -Interface $vnic0 -ManagementPortGroup $mgtNic

    if(!($dlr)) {
      Write-Host "Unable to create DLR $DLR_Name $_" -ForegroundColor "red"
    }
    else {
      $created++
    }
    $intRow++
  }
  $a = Release-Ref($WorkSheet_DLR)
  Write-Host "Added $created Distributed Logical Routers!" -ForegroundColor "green"
}




$stopwatch.Stop()
Write-Host "Elapsed time: ".$stopwatch.Elapsed

# Cleanup Excel object
$Excel.Quit()

$a = Release-Ref($WorkBook)
$a = Release-Ref($Excel)
