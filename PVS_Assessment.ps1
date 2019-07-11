# Carl Webster, CTP and Sr. Solutions Architect at Choice Solutions
# webster@carlwebster.com
# @carlwebster on Twitter
# http://www.CarlWebster.com
#
# Version 1.20 8-July-2019
#	Added to Farm properties, Citrix Provisioning license type: On-Premises or Cloud (new to 1808)
#	Added to vDisk properties, Accelerated Office Activation (new to 1906)
#	Added to vDisk properties, updated Write Cache types (new to 1811)
#		Private Image with Asynchronous IO
#		Cache on server, persistent with Asynchronous IO
#		Cache in device RAM with overflow on hard disk with Asynchronous IO

# Thanks to @jeffwouters for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(Mandatory=$False)] 
	[Alias("AA")]
	[string]$AdminAddress='',

	[parameter(Mandatory=$False)] 
	[switch]$CSV=$False,

	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[string]$Domain=$env:UserDnsDomain,

	[parameter(Mandatory=$False)] 
	[string]$User='',

	[parameter(Mandatory=$False)] 
	[SecureString]$Password='',

	[parameter(Mandatory=$False)] 
	[string]$Folder='',

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer='',

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From='',

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To=''
	
)

Set-StrictMode -Version Latest

if ($Folder -ne '')
{
	Write-Verbose -Message ('{0}: Testing folder path: {1}' -f (Get-Date), $Folder)
	# Check if $Folder exists
	if (Test-Path -Path $Folder -PathType Container) {
		#it exists and it is a folder
		Write-Verbose -Message ('{0}: Folder path $Folder exists and is a container' -f (Get-Date))
	} else {
		Write-Error "Folder $Folder does not exist or exists, but is not a folder.  Script cannot continue"
		Exit
	}
}

[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

$Script:AdvancedItems1         = New-Object System.Collections.ArrayList
$Script:AdvancedItems2         = New-Object System.Collections.ArrayList
$Script:ConfigWizItems         = New-Object System.Collections.ArrayList
$Script:BootstrapItems         = New-Object System.Collections.ArrayList
$Script:TaskOffloadItems       = New-Object System.Collections.ArrayList
$Script:PVSServiceItems        = New-Object System.Collections.ArrayList
$Script:VersionsToMerge        = New-Object System.Collections.ArrayList
$Script:NICIPAddresses 		   = @{}
$Script:StreamingIPAddresses   = New-Object System.Collections.ArrayList
$Script:BadIPs                 = New-Object System.Collections.ArrayList
$Script:EmptyDeviceCollections = New-Object System.Collections.ArrayList
$Script:MiscRegistryItems      = New-Object System.Collections.ArrayList
$Script:CacheOnServer          = New-Object System.Collections.ArrayList
$Script:MSHotfixes             = New-Object System.Collections.ArrayList	
$Script:WinInstalledComponents = New-Object System.Collections.ArrayList	
$Script:PVSProcessItems        = New-Object System.Collections.ArrayList

Function Get-IPAddress {
	#V1.16 added new function
	Param([string]$ComputerName)
	
	$IPAddress = "Unable to determine"
	
	Try {
		$IP = Test-Connection -ComputerName $ComputerName -Count 1 | Select-Object IPV4Address
	}
	Catch {
		$IP = "Unable to resolve IP address"
	}

	if ($? -and $Null -ne $IP -and $IP -ne "Unable to resolve IP address") {
		$IPAddress = $IP.IPV4Address.IPAddressToString
	}
	
	Return $IPAddress
}

Function validObject( [object] $object, [string] $topLevel ) {
	#function created 8-jan-2014 by Michael B. Smith
	if ($object) {
		if ((Get-Member -Name $topLevel -InputObject $object)) {
			Return $True
		}
	}
	Return $False
}

#region code for -hardware switch
Function GetComputerWMIInfo {
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 2-Apr-2018 to add ComputerOS information

	#Get Computer info
	Write-Verbose -Message ('{0}: Processing WMI Computer information' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Hardware information' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('Computer Information: {0}' -f $RemoteComputerName)
	Write-Verbose -Message 'General Computer'
	
	Try {
		$Results = Get-WmiObject -ComputerName $RemoteComputerName WIN32_ComputerSystem
	}
	Catch {
		$Results = $Null
	}
	
	if ($? -and $Null -ne $Results) {
		$ComputerItems = $Results | Select-Object -ExpandProperty Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		[string]$ComputerOS = (Get-WmiObject -class Win32_OperatingSystem -ComputerName $RemoteComputerName -EA 0).Caption

		ForEach($Item in $ComputerItems) {
			OutputComputerItem $Item $ComputerOS
	 	}
	} elseif (!$?) {
		Write-Verbose -Message ('{0}: Get-WmiObject win32_ComputerSystem failed for {1}' -f (Get-Date -UFormat "%F %r (%Z)"), $RemoteComputerName)
		Write-Warning ('Get-WmiObject win32_ComputerSystem failed for {1}' -f $RemoteComputerName)
		Line 2 ('Get-WmiObject win32_ComputerSystem failed for {1}' -f $RemoteComputerName)
		Line 2 ('On {1} you may need to run winmgmt /verifyrepository' -f $RemoteComputerName)
		Line 2 'and winmgmt /salvagerepository.  If this is a trusted Forest, you may'
		Line 2 'need to rerun the script with Domain Admin credentials from the trusted Forest.'
		Line 2 ''
	} else {
		Write-Verbose -Message ('{0}: No results Returned for Computer information' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output -Message "No results Returned for Computer information"
	}
	
	#Get Disk info
	Write-Verbose -Message ('{0}: Drive information' -f (Get-Date -UFormat "%F %r (%Z)"))

	Write-Output -Message "Drive(s)"

	Try
	{
		$Results = Get-WmiObject -ComputerName $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	if ($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			if ($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	} elseif (!$?) {
		Write-Verbose -Message ('{0}: Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Output -Message "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Output -Message "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Write-Output -Message "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Write-Output -Message "need to rerun the script with Domain Admin credentials from the trusted Forest."
	} else {
		Write-Verbose -Message ('{0}: No results Returned for Drive information' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output -Message "No results Returned for Drive information"
	}
	

	#Get CPU's and stepping
	Write-Verbose -Message ('{0}: Processor information' -f (Get-Date -UFormat "%F %r (%Z)"))

	Write-Output -Message "Processor(s)"

	Try
	{
		$Results = Get-WmiObject -ComputerName $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	if ($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	} elseif (!$?) {
		Write-Verbose -Message ('{0}: Get-WmiObject win32_Processor failed for $($RemoteComputerName)' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Output -Message "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Output -Message "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Write-Output -Message "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Write-Output -Message "need to rerun the script with Domain Admin credentials from the trusted Forest."
	} else {
		Write-Verbose -Message ('{0}: No results Returned for Processor information' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output -Message "No results Returned for Processor information"
	}

	#Get Nics
	Write-Verbose -Message ('{0}: NIC information' -f (Get-Date -UFormat "%F %r (%Z)"))

	Write-Output -Message "Network Interface(s)"

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -ComputerName $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch {
		$Results = $Null
	}

	if ($? -and $Null -ne $Results)	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		if ($Null -eq $Nics) {
			$GotNics = $False 
		} else { 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		if ($GotNics) {
			ForEach($nic in $nics) {
				Try {
					$ThisNic = Get-WmiObject -ComputerName $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				Catch {
					$ThisNic = $Null
				}
				
				if ($? -and $Null -ne $ThisNic) {
					OutputNicItem $Nic $ThisNic $RemoteComputerName
				} elseif (!$?) {
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose -Message ('{0}: Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)' -f (Get-Date -UFormat "%F %r (%Z)"))
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Output -Message "Error retrieving NIC information"
					Write-Output -Message "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Output -Message "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
					Write-Output -Message "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
					Write-Output -Message "need to rerun the script with Domain Admin credentials from the trusted Forest."
				} else {
					Write-Verbose -Message ('{0}: No results Returned for NIC information' -f (Get-Date -UFormat "%F %r (%Z)"))
					Line 2 'No results Returned for NIC information'
				}
			}
		}	
	} elseif (!$?) {
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose -Message ('{0}: Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Output -Message "Error retrieving NIC configuration information"
		Write-Output -Message "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Output -Message "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Write-Output -Message "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Write-Output -Message "need to rerun the script with Domain Admin credentials from the trusted Forest."
	} else {
		Write-Verbose -Message ('{0}: No results Returned for NIC configuration information' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output -Message "No results Returned for NIC configuration information"
	}
	
	Line 0 ''
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS)
	# modified 2-Apr-2018 to add Operating System information
	
	Line 2 "Manufacturer: " $Item.manufacturer
	Line 2 "Model: " $Item.model
	Line 2 "Domain: " $Item.domain
	Line 2 "Operating System: " $OS
	Write-Output -Message "Total Ram: $($Item.totalphysicalram) GB"
	Line 2 "Physical Processors (sockets): " $Item.NumberOfProcessors
	Line 2 "Logical Processors (cores w/HT): " $Item.NumberOfLogicalProcessors
	Line 2 ''
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ''
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ''
	if (![String]::IsNullOrEmpty($drive.volumedirty))
	{
		if ($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		} else {
			$xVolumeDirty = "No"
		}
	}

	Line 3 "Caption: " $drive.caption
	Write-Output -Message "Size: $($drive.drivesize) GB"
	if (![String]::IsNullOrEmpty($drive.filesystem))
	{
		Line 3 "File System: " $drive.filesystem
	}
	Write-Output -Message "Free Space: $($drive.drivefreespace) GB"
	if (![String]::IsNullOrEmpty($drive.volumename))
	{
		Line 3 "Volume Name: " $drive.volumename
	}
	if (![String]::IsNullOrEmpty($drive.volumedirty))
	{
		Line 3 "Volume is Dirty: " $xVolumeDirty
	}
	if (![String]::IsNullOrEmpty($drive.volumeserialnumber))
	{
		Line 3 "Volume Serial #: " $drive.volumeserialnumber
	}
	Line 3 "Drive Type: " $xDriveType
	Line 3 ''
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ''
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break }
		2	{$xAvailability = "Unknown"; Break }
		3	{$xAvailability = "Running or Full Power"; Break }
		4	{$xAvailability = "Warning"; Break }
		5	{$xAvailability = "In Test"; Break }
		6	{$xAvailability = "Not Applicable"; Break }
		7	{$xAvailability = "Power Off"; Break }
		8	{$xAvailability = "Off Line"; Break }
		9	{$xAvailability = "Off Duty"; Break }
		10	{$xAvailability = "Degraded"; Break }
		11	{$xAvailability = "Not Installed"; Break }
		12	{$xAvailability = "Install Error"; Break }
		13	{$xAvailability = "Power Save - Unknown"; Break }
		14	{$xAvailability = "Power Save - Low Power Mode"; Break }
		15	{$xAvailability = "Power Save - Standby"; Break }
		16	{$xAvailability = "Power Cycle"; Break }
		17	{$xAvailability = "Power Save - Warning"; Break }
		Default	{$xAvailability = "Unknown"; Break }
	}

	Line 3 "Name: " $processor.name
	Line 3 "Description: " $processor.description
	Write-Output -Message "Max Clock Speed: $($processor.maxclockspeed) MHz"
	if ($processor.l2cachesize -gt 0)
	{
		Write-Output -Message "L2 Cache Size: $($processor.l2cachesize) KB"
	}
	if ($processor.l3cachesize -gt 0)
	{
		Write-Output -Message "L3 Cache Size: $($processor.l3cachesize) KB"
	}
	if ($processor.numberofcores -gt 0)
	{
		Line 3 "# of Cores: " $processor.numberofcores
	}
	if ($processor.numberoflogicalprocessors -gt 0)
	{
		Line 3 "# of Logical Procs (cores w/HT): " $processor.numberoflogicalprocessors
	}
	Line 3 "Availability: " $xAvailability
	Line 3 ''
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string] $ComputerName)
	
	#V1.16 change how $powerMgmt is retrieved
	if (validObject $ThisNic PowerManagementSupported)
	{
		$powerMgmt = $ThisNic.PowerManagementSupported
	}
	
	if ($powerMgmt)
	{
		$powerMgmt = Get-WmiObject -ComputerName $ComputerName MSPower_DeviceEnable -Namespace root\wmi | Where-Object {$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

		if ($? -and $Null -ne $powerMgmt)
		{
			if ($powerMgmt.Enable -eq $True)
			{
				$PowerSaving = "Enabled"
			} else {
				$PowerSaving = "Disabled"
			}
		} else {
			$PowerSaving = "N/A"
		}
	} else {
		$PowerSaving = "Not Supported"
	}

	
	$xAvailability = ''
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break }
		2	{$xAvailability = "Unknown"; Break }
		3	{$xAvailability = "Running or Full Power"; Break }
		4	{$xAvailability = "Warning"; Break }
		5	{$xAvailability = "In Test"; Break }
		6	{$xAvailability = "Not Applicable"; Break }
		7	{$xAvailability = "Power Off"; Break }
		8	{$xAvailability = "Off Line"; Break }
		9	{$xAvailability = "Off Duty"; Break }
		10	{$xAvailability = "Degraded"; Break }
		11	{$xAvailability = "Not Installed"; Break }
		12	{$xAvailability = "Install Error"; Break }
		13	{$xAvailability = "Power Save - Unknown"; Break }
		14	{$xAvailability = "Power Save - Low Power Mode"; Break }
		15	{$xAvailability = "Power Save - Standby"; Break }
		16	{$xAvailability = "Power Cycle"; Break }
		17	{$xAvailability = "Power Save - Warning"; Break }
		Default	{$xAvailability = "Unknown"; Break }
	}

	$xIPAddresses = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddresses += "$($IPAddress)"
		#$Script:NICIPAddresses.Add($ComputerName, $IPAddress)
		if ($Script:NICIPAddresses.ContainsKey($ComputerName)) 
		{
			$MultiIP = @()
			$MultiIP += $Script:NICIPAddresses.Item($ComputerName)
			$MultiIP += $IPAddress
			$Script:NICIPAddresses.Item($ComputerName) = $MultiIP
		} else {
			$Script:NICIPAddresses.Add($ComputerName,$IPAddress)
		}
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	if ($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	if ($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ''
	if ($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	} else {
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ''
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ''
	if ($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	} else {
		$xwinsenablelmhostslookup = "No"
	}

	Line 3 "Name: " $ThisNic.Name
	if ($ThisNic.Name -ne $nic.description)
	{
		Line 3 "Description: " $nic.description
	}
	Line 3 "Connection ID: " $ThisNic.NetConnectionID
	Line 3 "Manufacturer: " $ThisNic.manufacturer
	Line 3 "Availability: " $xAvailability
    Line 3 "Allow the computer to turn off this device to save power: " $PowerSaving
	Line 3 "Physical Address: " $nic.macaddress
	Line 3 "IP Address: " $xIPAddresses[0]
	$cnt = -1
	ForEach($tmp in $xIPAddresses)
	{
		$cnt++
		if ($cnt -gt 0)
		{
			Line 4 "    " $tmp
		}
	}
	Line 3 "Default Gateway: " $Nic.Defaultipgateway
	Line 3 "Subnet Mask: " $xIPSubnet[0]
	$cnt = -1
	ForEach($tmp in $xIPSubnet)
	{
		$cnt++
		if ($cnt -gt 0)
		{
			Line 4 "     " $tmp
		}
	}
	if ($nic.dhcpenabled)
	{
		$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
		$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
		Line 3 "DHCP Enabled: " $nic.dhcpenabled
		Line 3 "DHCP Lease Obtained: " $dhcpleaseobtaineddate
		Line 3 "DHCP Lease Expires: " $dhcpleaseexpiresdate
		Line 3 "DHCP Server:" $nic.dhcpserver
	}
	if (![String]::IsNullOrEmpty($nic.dnsdomain))
	{
		Line 3 "DNS Domain: " $nic.dnsdomain
	}
	if ($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		[int]$x = 1
		Line 3 "DNS Search Suffixes: " $xnicdnsdomainsuffixsearchorder[0]
		$cnt = -1
		ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
		{
			$cnt++
			if ($cnt -gt 0)
			{
				Line 4 "    " $tmp
			}
		}
	}
	Line 3 "DNS WINS Enabled: " $xdnsenabledforwinsresolution
	if ($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		[int]$x = 1
		Line 3 "DNS Servers: " $xnicdnsserversearchorder[0]
		$cnt = -1
		ForEach($tmp in $xnicdnsserversearchorder)
		{
			$cnt++
			if ($cnt -gt 0)
			{
				Line 4 "     " $tmp
			}
		}
	}
	Line 3 "NetBIOS Setting: " $xTcpipNetbiosOptions
	Line 3 "Enabled LMHosts: " $xwinsenablelmhostslookup
	if (![String]::IsNullOrEmpty($nic.winshostlookupfile))
	{
		Line 3 "Host Lookup File: " $nic.winshostlookupfile
	}
	if (![String]::IsNullOrEmpty($nic.winsprimaryserver))
	{
		Line 3 "Primary Server: " $nic.winsprimaryserver
	}
	if (![String]::IsNullOrEmpty($nic.winssecondaryserver))
	{
		Line 3 "Secondary Server: " $nic.winssecondaryserver
	}
	if (![String]::IsNullOrEmpty($nic.winsscopeid))
	{
		Line 3 "Scope ID: " $nic.winsscopeid
	}
	Line 0 ''
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose -Message ('{0}: Prepare to email' -f (Get-Date -UFormat "%F %r (%Z)"))
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	$error.Clear()
	if ($UseSSL)
	{
		Write-Verbose -Message ('{0}: Trying to send email using current user's credentials with SSL' -f (Get-Date -UFormat "%F %r (%Z)"))
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	} else {
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	if ($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose -Message ('{0}: Current user's credentials failed. Ask for usable credentials.' -f (Get-Date -UFormat "%F %r (%Z)"))

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		$error.Clear()
		if ($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		} else {
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		if ($? -and $Null -eq $e)
		{
			Write-Verbose -Message ('{0}: Email successfully sent using new credentials' -f (Get-Date -UFormat "%F %r (%Z)"))
		} else {
			Write-Verbose -Message ('{0}: Email was not sent:' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	} else {
		Write-Verbose -Message ('{0}: Email was not sent:' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

Function GetConfigWizardInfo
{
	Param([string]$ComputerName)
	
	$DHCPServicesValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "DHCPType" $ComputerName
	$PXEServiceValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "PXEType" $ComputerName
	
	$DHCPServices = ''
	$PXEServices = ''

	Switch ($DHCPServicesValue)
	{
		1073741824 {$DHCPServices = "The service that runs on another computer"; Break}
		0 {$DHCPServices = "Microsoft DHCP"; Break}
		1 {$DHCPServices = "Provisioning Services BOOTP service"; Break}
		2 {$DHCPServices = "Other BOOTP or DHCP service"; Break}
		Default {$DHCPServices = "Unable to determine DHCPServices: $($DHCPServicesValue)"; Break}
	}

	if ($DHCPServicesValue -eq 1073741824)
	{
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	} elseif ($DHCPServicesValue -eq 0) {
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Microsoft DHCP"; Break}
			1 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	} elseif ($DHCPServicesValue -eq 1) {
		$PXEServices = "N/A"
	} elseif ($DHCPServicesValue -eq 2) {
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	}

	$UserAccount1Value = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "Account1" $ComputerName
	$UserAccount3Value = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "Account3" $ComputerName
	
	$UserAccount = ''
	
	if ([String]::IsNullOrEmpty($UserAccount1Value) -and $UserAccount3Value -eq 1)
	{
		$UserAccount = "NetWork Service"
	} elseif ([String]::IsNullOrEmpty($UserAccount1Value) -and $UserAccount3Value -eq 0) {
		$UserAccount = "Local system account"
	} elseif (![String]::IsNullOrEmpty($UserAccount1Value)) {
		$UserAccount = $UserAccount1Value
	}

	$TFTPOptionValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "TFTPSetting" $ComputerName
	$TFTPOption = ''
	
	if ($TFTPOptionValue -eq 1)
	{
		$TFTPOption = "Yes"
		$TFTPBootstrapLocation = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Admin" "Bootstrap" $ComputerName
	} else {
		$TFTPOption = "No"
	}

	Write-Verbose -Message ('{0}: Gather Config Wizard info for Appendix C' -f (Get-Date -UFormat "%F %r (%Z)"))
	$obj1 = [PSCustomObject] @{
		ServerName        = $ComputerName
		DHCPServicesValue = $DHCPServicesValue
		PXEServicesValue  = $PXEServiceValue
		UserAccount       = $UserAccount
		TFTPOptionValue   = $TFTPOptionValue
	}
	$null = $Script:ConfigWizItems.Add($obj1)
	
	Write-Output -Message "Configuration Wizard Settings"
	Line 3 "DHCP Services: " $DHCPServices
	Line 3 "PXE Services: " $PXEServices
	Line 3 "User account: " $UserAccount
	Line 3 "TFTP Option: " $TFTPOption
	if ($TFTPOptionValue -eq 1)
	{
		Line 3 "TFTP Bootstrap Location: " $TFTPBootstrapLocation
	}
	
	Line 0 ''
}

Function GetDisableTaskOffloadInfo
{
	Param([string]$ComputerName)
	
	Write-Verbose -Message ('{0}:Gather TaskOffload info for Appendix E' -f (Get-Date -UFormat "%F %r (%Z)"))
	$TaskOffloadValue = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\TCPIP\Parameters" "DisableTaskOffload" $ComputerName
	
	if ($Null -eq $TaskOffloadValue)
	{
		$TaskOffloadValue = "Missing"
	}
	
	$obj1 = [PSCustomObject] @{
		ServerName       = $ComputerName	
		TaskOffloadValue = $TaskOffloadValue	
	}
	$null = $Script:TaskOffloadItems.Add($obj1)
	
	Write-Output -Message "TaskOffload Settings"
	Line 3 "Value: " $TaskOffloadValue
	
	Line 0 ''
}

Function Get-RegKeyToObject 
{
	#function contributed by Andrew Williamson @ Fujitsu Services
    param([string]$RegPath,
    [string]$RegKey,
    [string]$ComputerName)
	
    $val = Get-RegistryValue $RegPath $RegKey $ComputerName
	
    if ($Null -eq $val) 
	{
        $tmp = "Not set"
    } else {
	    $tmp = $val
    }
	
	$obj1 = [PSCustomObject] @{
		ServerName = $ComputerName	
		RegKey     = $RegPath	
		RegValue   = $RegKey	
		Value      = $tmp	
	}
	$null = $Script:MiscRegistryItems.Add($obj1)
}

Function GetMiscRegistryKeys
{
	Param([string]$ComputerName)
	
	#look for the following registry keys and values on PVS servers
		
	#Registry Key                                                      Registry Value                 
	#=================================================================================================
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        AutoUpdateUserCache            
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        LoggingLevel 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        SkipBootMenu                   
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        UseManagementIpInCatalog       
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        UseTemplateBootOrder           
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    IPv4Address                    
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    PortBase 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    PortCount 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Manager                GeneralInetAddr                
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\SoapServer             PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipBootMenu                   
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipRIMS                       
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipRIMSforPrivate             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcHDNoIntermediateBuffering    
	#HKLM:\SYSTEM\CurrentControlSet\services\BNIStack\Parameters       WcRamConfiguration             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcWarningIncrement             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcWarningPercent               
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNNS\Parameters           EnableOffload                  
	#HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters         InitTimeoutSec           
	#HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters         MaxBindRetry             
	#HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters         InitTimeoutSec           
	#HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters         MaxBindRetry      
	
	Write-Verbose -Message ('{0}: Gather Misc Registry Key data for Appendix K' -f (Get-Date -UFormat "%F %r (%Z)"))

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "AutoUpdateUserCache" $ComputerName

	#https://support.citrix.com/article/CTX135299
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "SkipBootMenu" $ComputerName

	#https://support.citrix.com/article/CTX142613
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "UseManagementIpInCatalog" $ComputerName

	#https://support.citrix.com/article/CTX142613
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX200196
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX200196
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Manager" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX135299
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "UseTemplateBootOrder" $ComputerName

	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "SkipRIMS" $ComputerName

	#https://support.citrix.com/article/CTX200233
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "SkipRIMSforPrivate" $ComputerName

	#https://support.citrix.com/article/CTX126042
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcHDNoIntermediateBuffering" $ComputerName

	#https://support.citrix.com/article/CTX139849
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\services\BNIStack\Parameters" "WcRamConfiguration" $ComputerName

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcWarningIncrement" $ComputerName

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcWarningPercent" $ComputerName

	#https://support.citrix.com/article/CTX117374
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNNS\Parameters" "EnableOffload" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters" "InitTimeoutSec" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters" "MaxBindRetry" $ComputerName

	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters" "InitTimeoutSec" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters" "MaxBindRetry" $ComputerName

	#regkeys recommended by Andrew Williamson @ Fujitsu Services
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "LoggingLevel" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "PortBase" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "PortCount" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\SoapServer" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "PortOffset" $ComputerName
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	if ($ComputerName -eq $env:ComputerName)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		if ($key)
		{
			Return $key.GetValue($name, $Null)
		} else {
			Return $Null
		}
	} else {
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey = $Reg.OpenSubKey($path2)
		If ($RegKey)
		{
			$Results = $RegKey.GetValue($name)

			if ($Null -ne $Results)
			{
				Return $Results
			} else {
				Return $Null
			}
		} else {
			Return $Null
		}
	}
}

Function BuildPVSObject
{
	Param([string]$MCLIGetWhat = '', [string]$MCLIGetParameters = '', [string]$TextForErrorMsg = '')

	$error.Clear()

	if ($MCLIGetParameters -ne '')
	{
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)" -p "$($MCLIGetParameters)"
	} else {
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)"
	}

	if ($error.Count -eq 0)
	{
		$PluralObject = @()
		$SingleObject = $Null
		ForEach($record in $MCLIGetResult)
		{
			if ($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
			{
				if ($Null -ne $SingleObject)
				{
					$PluralObject += $SingleObject
				}
				$SingleObject = new-object System.Object
			}

			$index = $record.IndexOf(':')
			if ($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value    = $record.SubString($index + 2)
				if ($property -ne "Executing")
				{
					Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
				}
			}
		}
		$PluralObject += $SingleObject
		Return $PluralObject
	} else {
		Write-Output -Message "$($TextForErrorMsg) could not be retrieved"
		Line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	}
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | ForEach-Object {$_.name}
	$registeredSnapins += get-pssnapin -Registered | ForEach-Object {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		if (!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			if (!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				if (!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			} else {
				#Snapin is registered, but not loaded, loading it now:
				Write-Output -Message "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	if ($FoundMissingSnapin) {
		Write-Warning -Message 'Missing Windows PowerShell snap-ins Detected:'
		$missingSnapins | ForEach-Object { Write-Warning -Message "($_)" }
		return $False
	} else {
		Return $True
	}
}

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( '' )
		$tabs--
	}

	if ($nonewline ) {
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( $name + $value )
	} else {
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.AppendLine( $name + $value )
	}
}
	
Function SaveandCloseTextDocument {
	# RFE?: SHould this function use Add-Content or Set-Content instead?
	if ($Host.Version.CompareTo( [System.Version]'2.0' ) -eq 0 ) {
		Write-Verbose -Message ('{0}: Saving for PowerShell Version 2' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output $global:Output.ToString() | Out-File $Script:Filename1 2>$Null
	} else {
		Write-Verbose -Message ('{0}: Saving for PowerShell Version 3 or later' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output $global:Output.ToString() | Out-File $Script:Filename1 4>$Null
	}
}

Function SetFileName1 {
	Param([string]$OutputFileName)
	if ($Folder -eq '') {
		$Script:pwdpath = $pwd.Path
	} else {
		$Script:pwdpath = $Folder
	}

	# if ($Script:pwdpath.EndsWith("\")) {
	# 	#remove the trailing \
	# 	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
	# }

	# [string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).txt"
	[string]$Script:FileName1 = Join-Path -Path $Script:pwdpath -ChildPath ('{0}.txt' -f $OutputFileName)
}

Function ElevatedSession
{
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	if ($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose -Message ('{0}: This is an elevated PowerShell session' -f (Get-Date -UFormat "%F %r (%Z)"))
		Return $True
	} else {
		Write-Output -Message ('{0}: This is NOT an elevated PowerShell session' -f (Get-Date -UFormat "%F %r (%Z)"))
		Return $False
	}
}

Function SetupRemoting {
	# setup remoting if $AdminAddress is not empty
	[bool]$Script:Remoting = $False
	if (![System.String]::IsNullOrEmpty($AdminAddress)) {
		#since we are setting up remoting, the script must be run from an elevated PowerShell session
		$Elevated = ElevatedSession

		if (-not $Elevated ) {
			Write-Warning -Message 'Remoting to another PVS server was requested but this is not an elevated PowerShell session.'
			Write-Warning -Message 'Using -AdminAddress requires the script be run from an elevated PowerShell session.'
			Write-Warning -Message 'Please run the script from an elevated PowerShell session. Script cannot continue'
			Exit
		} else {
			Write-Verbose -Message 'This is an elevated PowerShell session.'
		}
		
		if (![System.String]::IsNullOrEmpty($User))
		{
			if ([System.String]::IsNullOrEmpty($Domain))
			{
				$Domain = Read-Host "Domain name for user is required. Enter Domain name for user"
			}		

			if ([System.String]::IsNullOrEmpty($Password))
			{
				$Password = Read-Host "Password for user is required. Enter password for user"
			}		
			$error.Clear()
			mcli-run SetupConnection -p server="$($AdminAddress)",user="$($User)",domain="$($Domain)",password="$($Password)"
		} else {
			$error.Clear()
			mcli-run SetupConnection -p server="$($AdminAddress)"
		}

		if ($error.Count -eq 0)
		{
			$Script:Remoting = $True
			Write-Verbose -Message ('{0}: This script is being run remotely against server $($AdminAddress)' -f (Get-Date -UFormat "%F %r (%Z)"))
			if (![System.String]::IsNullOrEmpty($User))
			{
				Write-Verbose -Message ('{0}: User=$($User)' -f (Get-Date -UFormat "%F %r (%Z)"))
				Write-Verbose -Message ('{0}: Domain=$($Domain)' -f (Get-Date -UFormat "%F %r (%Z)"))
			}
		} else {
			Write-Warning "Remoting could not be setup to server $($AdminAddress)"
			Write-Warning "Error returned is " $error[0]
			Write-Warning "Script cannot continue"
			Exit
		}
	} else {
		#added V1.17
		#if $AdminAddress is '', get actual server name
		if ($AdminAddress -eq '')
		{
			$Script:AdminAddress = $env:ComputerName
		}
	}
}

Function VerifyPVSServices
{
	Write-Verbose -Message ('{0}: Verifying PVS SOAP and Stream Services are running' -f (Get-Date -UFormat "%F %r (%Z)"))
	$soapserver = $Null
	$StreamService = $Null

	if ($Script:Remoting)
	{
		$soapserver = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
		$StreamService = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
	} else {
		$soapserver = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
		$StreamService = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
	}

	if ($soapserver.Status -ne "Running")
	{
		if ($Script:Remoting)
		{
			Write-Warning "The Citrix PVS Soap Server service is not Started on server $($AdminAddress)"
		} else {
			Write-Warning "The Citrix PVS Soap Server service is not Started"
		}
		Write-Error "Script cannot continue.  See message above."
		Exit
	}

	if ($StreamService.Status -ne "Running")
	{
		if ($Script:Remoting)
		{
			Write-Warning "The Citrix PVS Stream Service service is not Started on server $($AdminAddress)"
		} else {
			Write-Warning "The Citrix PVS Stream Service service is not Started"
		}
		Write-Error "Script cannot continue.  See message above."
		Exit
	}
}

Function VerifyPVSSOAPService
{
	Param([string]$PVSServer='')
	
	Write-Verbose -Message ('{0}: Verifying server $($PVSServer) is online' -f (Get-Date -UFormat "%F %r (%Z)"))
	if (Test-Connection -ComputerName $server.servername -quiet -EA 0) {
		Write-Verbose -Message ('{0}: Verifying PVS SOAP Service is running on server $($PVSServer)' -f (Get-Date -UFormat "%F %r (%Z)"))
		$soapserver = $Null

		$soapserver = Get-Service -ComputerName $PVSServer -EA 0 | Where-Object {$_.Name -like "soapserver"}

		if ($soapserver.Status -ne "Running")
		{
			Write-Warning "The Citrix PVS Soap Server service is not Started on server $($PVSServer)"
			Write-Warning "Server $($PVSServer) cannot be processed.  See message above."
			Return $False
		} else {
			Return $True
		}
	} else {
		Write-Warning "The server $($PVSServer) is offLine or unreachable."
		Write-Warning "Server $($PVSServer) cannot be processed.  See message above."
		Return $False
	}
}

Function GetPVSVersion
{
	#get PVS major version
	Write-Verbose -Message ('{0}: Getting PVS version info' -f (Get-Date -UFormat "%F %r (%Z)"))

	$error.Clear()
	$tempversion = mcli-info version
	if ($? -and $error.Count -eq 0)
	{
		#build PVS version values
		$version = new-object System.Object 
		ForEach($record in $tempversion)
		{
			$index = $record.IndexOf(':')
			if ($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value = $record.SubString($index + 2)
				Add-Member -inputObject $version -MemberType NoteProperty -Name $property -Value $value
			}
		}
	} else {
		Write-Warning "PVS version information could not be retrieved"
		[int]$NumErrors = $Error.Count
		For($x=0; $x -le $NumErrors; $x++)
		{
			Write-Warning "Error(s) returned: " $error[$x]
		}
		Write-Error "Script is terminating"
		#without version info, script should not proceed
		Exit
	}

	$Script:PVSVersion     = $Version.mapiVersion.SubString(0,1)
	$Script:PVSFullVersion = $Version.mapiVersion
}

Function GetPVSFarm
{
	#build PVS farm values
	Write-Verbose -Message ('{0}: Build PVS farm values' -f (Get-Date -UFormat "%F %r (%Z)"))
	#there can only be one farm
	$GetWhat = "Farm"
	$GetParam = ''
	$ErrorTxt = "PVS Farm information"
	$Script:Farm = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	if ($Null -eq $Script:Farm)
	{
		#without farm info, script should not proceed
		Write-Error "PVS Farm information could not be retrieved.  Script is terminating."
		Exit
	}

	[string]$Script:Title = "PVS Assessment Report for Farm $($Script:farm.FarmName)"
	SetFileName1 "$($Script:farm.FarmName)_Assessment" #V1.16 add _Assessment
}

Function ProcessPVSFarm
{
	Write-Verbose -Message ('{0}: Processing PVS Farm Information' -f (Get-Date -UFormat "%F %r (%Z)"))

	$LicenseServerIPAddress = Get-IPAddress $Script:farm.licenseServer #added in V1.16
	
	#V1.17 see if the database server names contain an instance name. If so, remove it
	#V1.18 add test for port number - bug found by Johan Parlevliet 
	#V1.18 see if the database server names contain a port number. If so, remove it
	#V1.18 optimized code supplied by MBS
	$dbServer = $Script:farm.databaseServerName
	if (( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Verbose -Message ('{0}: Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address' -f (Get-Date -UFormat "%F %r (%Z)"))
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Verbose -Message ('{0}: dbServer now '$dbServer'' -f (Get-Date -UFormat "%F %r (%Z)"))
	}
	$SQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	$dbServer = $Script:farm.failoverPartnerServerName
	if (( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Verbose -Message ('{0}: Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address' -f (Get-Date -UFormat "%F %r (%Z)"))
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Verbose -Message ('{0}: dbServer now '$dbServer'' -f (Get-Date -UFormat "%F %r (%Z)"))
	}
	$FailoverSQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	#general tab
	Line 0 "PVS Farm Name: " $Script:farm.farmName
	Line 0 "Version: " $Script:PVSFullVersion
	
	Write-Verbose -Message ('{0}: Processing Licensing Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
	Line 0 "License server name: " $Script:farm.licenseServer
	Line 0 "License server IP: " $LicenseServerIPAddress
	Line 0 "License server port: " $Script:farm.licenseServerPort
	if ($Script:PVSVersion -eq "5")
	{
		Line 0 "Use Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
		if ($farm.licenseTradeUp -eq "1")
		{
			Write-Output -Message "Yes"
		} else {
			Write-Output -Message "No"
		}
	}
	if ($Script:PVSFullVersion -ge "7.19")
	{
		Line 0 "Citrix Provisioning license type" ''
		if ($farm.LicenseSKU -eq 2)
		{
			Write-Output -Message "On-Premises: " "Yes"
			Line 2 "Use Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
			if ($farm.licenseTradeUp -eq "1")
			{
				Write-Output -Message "Yes"
			} else {
				Write-Output -Message "No"
			}
			Write-Output -Message "Cloud: " "No"
		} else {
			Write-Output -Message "On-Premises: " "No"
			Write-Output -Message "Use Datacenter licenses for desktops if no Desktop licenses are available: No"
			Write-Output -Message "Cloud: " "Yes"
		}
	} elseif ($Script:PVSFullVersion -ge "7.13") {
		Line 1 "Use Datacenter licenses for desktops if no Desktop licenses are available: " $DatacenterLicense
	}

	Line 0 "Enable auto-add: " -nonewline
	if ($farm.autoAddEnabled -eq "1")
	{
		Write-Output -Message "Yes"
		Line 0 "Add new devices to this site: " $farm.DefaultSiteName
		$Script:FarmAutoAddEnabled = $True
	} else {
		Line 0 "No"	
		$Script:FarmAutoAddEnabled = $False
	}	
	
	#options tab
	Write-Verbose -Message ('{0}: Processing Options Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
	Line 0 "Enable auditing: " -nonewline
	if ($Script:farm.auditingEnabled -eq "1")
	{
		Write-Output -Message "Yes"
	} else {
		Write-Output -Message "No"
	}
	Line 0 "Enable offline database support: " -nonewline
	if ($Script:farm.offlineDatabaseSupportEnabled -eq "1")
	{
		Line 0 "Yes"	
	} else {
		Write-Output -Message "No"
	}

	if ($Script:PVSVersion -eq "6" -or $Script:PVSVersion -eq "7")
	{
		#vDisk Version tab
		Write-Verbose -Message ('{0}: Processing vDisk Version Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
		Write-Output -Message "vDisk Version"
		Line 1 "Alert if number of versions from base image exceeds: " $Script:farm.maxVersions
		Line 1 "Default access mode for new merge versions: " -nonewline
		Switch ($Script:farm.mergeMode)
		{
			0   {Line 0 "Production"; Break }
			1   {Line 0 "Test"; Break }
			2   {Line 0 "Maintenance"; Break}
			Default {Line 0 "Default access mode could not be determined: $($Script:farm.mergeMode)"; Break}
		}
	}
	
	#status tab
	Write-Verbose -Message ('{0}: Processing Status Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
	Line 0 "Database server: " $Script:farm.databaseServerName
	Line 0 "Database server IP: " $SQLServerIPAddress
	Line 0 "Database instance: " $Script:farm.databaseInstanceName
	Line 0 "Database: " $Script:farm.databaseName
	Line 0 "Failover Partner Server: " $Script:farm.failoverPartnerServerName
	Line 0 "Failover Partner Server IP: " $FailoverSQLServerIPAddress
	Line 0 "Failover Partner Instance: " $Script:farm.failoverPartnerInstanceName
	if ($Script:farm.adGroupsEnabled -eq "1")
	{
		Write-Output -Message "Active Directory groups are used for access rights"
	} else {
		Write-Output -Message "Active Directory groups are not used for access rights"
	}
	Line 0 ''
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function ProcessPVSSite
{
	#build site values
	Write-Verbose -Message ('{0}: Processing Sites' -f (Get-Date -UFormat "%F %r (%Z)"))
	$GetWhat = "site"
	$GetParam = ''
	$ErrorTxt = "PVS Site information"
	$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	if ($Null -eq $PVSSites)
	{
		Write-Warning -Message "$(Get-Date): No Sites Found"
		Write-Output -Message "No Sites Found "
	} else {
		ForEach($PVSSite in $PVSSites)
		{
			Write-Verbose -Message ('{0}: Processing Site $($PVSSite.siteName)' -f (Get-Date -UFormat "%F %r (%Z)"))
			Line 0 "Site Name: " $PVSSite.siteName

			#security tab
			Write-Verbose -Message ('{0}: Processing Security Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			$temp = $PVSSite.SiteName
			$GetWhat = "authgroup"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Groups with Site Administrator access"
			$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			if ($Null -ne $authGroups)
			{
				Write-Output -Message "Groups with Site Administrator access:"
				ForEach($Group in $authgroups)
				{
					Line 2 $Group.authGroupName
				}
			} else {
				Write-Output -Message "Groups with Site Administrator access: No Site Administrators defined"
			}

			#MAK tab
			#MAK User and Password are encrypted

			# options tab
			Write-Verbose -Message ('{0}: Processing Options Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			if ($PVSVersion -eq "5" -or (($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $FarmAutoAddEnabled)) {
				Write-Output -Message "Add new devices to this collection: " -nonewline
				if ($PVSSite.DefaultCollectionName) {
					Line 0 $PVSSite.DefaultCollectionName
				} else {
					Write-Output -Message "<No Default collection>"
				}
			}
			if ($PVSVersion -eq "6" -or $PVSVersion -eq "7") {
				if ($PVSVersion -eq "6") {
					Write-Output -Message "Seconds between vDisk inventory scans: " $PVSSite.inventoryFilePollingInterval
				}

				#vDisk Update
				Write-Verbose -Message ('{0}: Processing vDisk Update Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				if ($PVSSite.enableDiskUpdate -eq "1") {
					Write-Output -Message "Enable automatic vDisk updates on this site: Yes"
					Write-Output -Message "Server to run vDisk updates for this site: " $PVSSite.diskUpdateServerName
				} else {
					Write-Output -Message "Enable automatic vDisk updates on this site: No"
				}
			}
			Line 0 ''
			
			#process all servers in site
			Write-Verbose -Message ('{0}: Processing Servers in Site $($PVSSite.siteName)' -f (Get-Date -UFormat "%F %r (%Z)"))
			$temp = $PVSSite.SiteName
			$GetWhat = "server"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Servers for Site $temp"
			$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			if ($Null -eq $servers) {
				Write-Warning -Message "$(Get-Date): No Servers Found in Site $($PVSSite.siteName)"
				Write-Output -Message "No Servers Found in Site $($PVSSite.siteName)"
			} else {
				Write-Output -Message "Servers"
				ForEach($Server in $Servers) {
					#first make sure the SOAP service is running on the server
					if (VerifyPVSSOAPService $Server.serverName) {
						Write-Verbose -Message ('{0}: Processing Server $($Server.serverName)' -f (Get-Date -UFormat "%F %r (%Z)"))
						#general tab
						Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
						Line 2 "Name: " $Server.serverName
						Line 2 "Log events to the server's Windows Event Log: " -nonewline
						if ($Server.eventLoggingEnabled -eq "1")
						{
							Write-Output -Message "Yes"
						} else {
							Write-Output -Message "No"
						}
							
						Write-Verbose -Message ('{0}: Processing Network Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
						$test = $Server.ip.ToString()
						$test1 = $test.replace(",",", ")
						
						$tmparray= @($server.ip.split(","))
						
						ForEach($item in $tmparray)
						{
							$obj1 = [PSCustomObject] @{
								ServerName = $Server.serverName							
								IPAddress  = $item							
							}
							$null = $Script:StreamingIPAddresses.Add($obj1)
						}
						if ($Script:PVSVersion -eq "7")
						{
							Line 2 "Streaming IP addresses: " $test1
						} else {
							Line 2 "IP addresses: " $test1
						}
						Line 2 "First port: " $Server.firstPort
						Line 2 "Last port: " $Server.lastPort
						if ($Script:PVSVersion -eq "7")
						{
							Line 2 "Management IP: " $Server.managementIp
						}
							
						#create array for appendix A
						
						Write-Verbose -Message ('{0}: Gather Advanced server info for Appendix A and B' -f (Get-Date -UFormat "%F %r (%Z)"))
						$obj1 = [PSCustomObject] @{
							ServerName              = $Server.serverName						
							ThreadsPerPort          = $Server.threadsPerPort						
							BuffersPerThread        = $Server.buffersPerThread						
							ServerCacheTimeout      = $Server.serverCacheTimeout						
							LocalConcurrentIOLimit  = $Server.localConcurrentIoLimit						
							RemoteConcurrentIOLimit = $Server.remoteConcurrentIoLimit						
							maxTransmissionUnits    = $Server.maxTransmissionUnits						
							IOBurstSize             = $Server.ioBurstSize						
							NonBlockingIOEnabled    = $Server.nonBlockingIoEnabled						
						}
						$null = $Script:AdvancedItems1.Add($obj1)
						
						$obj2 = [PSCustomObject] @{
							ServerName              = $Server.serverName						
							BootPauseSeconds        = $Server.bootPauseSeconds						
							MaxBootSeconds          = $Server.maxBootSeconds						
							MaxBootDevicesAllowed   = $Server.maxBootDevicesAllowed						
							vDiskCreatePacing       = $Server.vDiskCreatePacing						
							LicenseTimeout          = $Server.licenseTimeout						
						}
						$null = $Script:AdvancedItems2.Add($obj2)
						
						GetComputerWMIInfo $server.ServerName
							
						GetConfigWizardInfo $server.ServerName
							
						GetDisableTaskOffloadInfo $server.ServerName
							
						GetBootstrapInfo $server
							
						GetPVSServiceInfo $server.ServerName

						GetBadStreamingIPAddresses $server.ServerName
						
						GetMiscRegistryKeys $server.ServerName
						
						GetMicrosoftHotfixes $server.ServerName
						
						GetInstalledRolesAndFeatures $server.ServerName
						
						GetPVSProcessInfo $server.ServerName
					} else {
						Line 2 "Name: " $Server.serverName
						Write-Output -Message "Server was not processed because the server was offLine or the SOAP Service was not running"
						Line 0 ''
					}
				}
			}

			#process all device collections in site
			Write-Verbose -Message ('{0}: Processing all device collections in site' -f (Get-Date -UFormat "%F %r (%Z)"))
			$Temp = $PVSSite.SiteName
			$GetWhat = "Collection"
			$GetParam = "siteName = $Temp"
			$ErrorTxt = "Device Collection information"
			$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			if ($Null -ne $Collections)
			{
				Write-Output -Message "Device Collections"
				ForEach($Collection in $Collections)
				{
					Write-Verbose -Message ('{0}: Processing Collection $($Collection.collectionName)' -f (Get-Date -UFormat "%F %r (%Z)"))
					Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
					Line 2 "Name: " $Collection.collectionName

					Write-Verbose -Message ('{0}: Processing Security Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
					$Temp = $Collection.collectionId
					$GetWhat = "authGroup"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Collection information"
					$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

					$DeviceAdmins = $False
					if ($Null -ne $AuthGroups)
					{
						Write-Output -Message "Groups with 'Device Administrator' access:"
						ForEach($AuthGroup in $AuthGroups)
						{
							$Temp = $authgroup.authGroupName
							$GetWhat = "authgroupusage"
							$GetParam = "authgroupname = $Temp"
							$ErrorTxt = "Device Collection Administrator usage information"
							$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
							if ($Null -ne $AuthGroupUsages)
							{
								ForEach($AuthGroupUsage in $AuthGroupUsages)
								{
									if ($AuthGroupUsage.role -eq "300")
									{
										$DeviceAdmins = $True
										Line 3 $authgroup.authGroupName
									}
								}
							}
						}
					}
					if (!$DeviceAdmins)
					{
						Write-Output -Message "Groups with 'Device Administrator' access: None defined"
					}

					$DeviceOperators = $False
					if ($Null -ne $AuthGroups)
					{
						Write-Output -Message "Groups with 'Device Operator' access:"
						ForEach($AuthGroup in $AuthGroups)
						{
							$Temp = $authgroup.authGroupName
							$GetWhat = "authgroupusage"
							$GetParam = "authgroupname = $Temp"
							$ErrorTxt = "Device Collection Operator usage information"
							$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
							if ($Null -ne $AuthGroupUsages)
							{
								ForEach($AuthGroupUsage in $AuthGroupUsages)
								{
									if ($AuthGroupUsage.role -eq "400")
									{
										$DeviceOperators = $True
										Line 3 $authgroup.authGroupName
									}
								}
							}
						}
					}
					if (!$DeviceOperators)
					{
						Write-Output -Message "Groups with 'Device Operator' access: None defined"
					}

					Write-Verbose -Message ('{0}: Processing Auto-Add Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
					if ($Script:FarmAutoAddEnabled)
					{
						Line 2 "Template target device: " $Collection.templateDeviceName
						if (![String]::IsNullOrEmpty($Collection.autoAddPrefix) -or ![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Write-Output -Message "Device Name"
						}
						if (![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Line 3 "Prefix: " $Collection.autoAddPrefix
						}
						Line 3 "Length: " $Collection.autoAddNumberLength
						Line 3 "Zero fill: " -nonewline
						if ($Collection.autoAddZeroFill -eq "1")
						{
							Write-Output -Message "Yes"
						} else {
							Write-Output -Message "No"
						}
						if (![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Line 3 "Suffix: " $Collection.autoAddSuffix
						}
						Line 3 "Last incremental #: " $Collection.lastAutoAddDeviceNumber
					} else {
						Write-Output -Message "The auto-add feature is not enabled at the PVS Farm level"
					}
					#for each collection process each device
					Write-Verbose -Message ('{0}: Processing the first device in each collection' -f (Get-Date -UFormat "%F %r (%Z)"))
					$Temp = $Collection.collectionId
					$GetWhat = "deviceInfo"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Info information"
					$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					
					if ($Null -ne $Devices)
					{
						Line 0 ''
						$Device = $Devices[0]
						Write-Verbose -Message ('{0}: Processing Device $($Device.deviceName)' -f (Get-Date -UFormat "%F %r (%Z)"))
						if ($Device.type -eq "3")
						{
							Write-Output -Message "Device with Personal vDisk Properties"
						} else {
							Write-Output -Message "Target Device Properties"
						}
						Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
						Line 3 "Name: " $Device.deviceName
						if (($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $Device.type -ne "3")
						{
							Line 3 "Type: " -nonewline
							Switch ($Device.type)
							{
								0 {Line 0 "Production"; Break}
								1 {Line 0 "Test"; Break}
								2 {Line 0 "Maintenance"; Break}
								3 {Line 0 "Personal vDisk"; Break}
								Default {Line 0 "Device type could not be determined: $($Device.type)"; Break}
							}
						}
						if ($Device.type -ne "3")
						{
							Line 3 "Boot from: " -nonewline
							Switch ($Device.bootFrom)
							{
								1 {Line 0 "vDisk"; Break}
								2 {Line 0 "Hard Disk"; Break}
								3 {Line 0 "Floppy Disk"; Break}
								Default {Line 0 "Boot from could not be determined: $($Device.bootFrom)"; Break}
							}
						}
						Line 3 "Port: " $Device.port
						if ($Device.type -ne "3")
						{
							Line 3 "Disabled: " -nonewline
							if ($Device.enabled -eq "1")
							{
								Write-Output -Message "No"
							} else {
								Write-Output -Message "Yes"
							}
						} else {
							Line 3 "vDisk: " $Device.diskLocatorName
							Line 3 "Personal vDisk Drive: " $Device.pvdDriveLetter
						}
						Write-Verbose -Message ('{0}: Processing vDisks Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
						#process all vdisks for this device
						$Temp = $Device.deviceName
						$GetWhat = "DiskInfo"
						$GetParam = "deviceName = $Temp"
						$ErrorTxt = "Device vDisk information"
						$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
						if ($Null -ne $vDisks)
						{
							ForEach($vDisk in $vDisks)
							{
								Write-Output -Message "vDisk Name: $($vDisk.storeName)`\$($vDisk.diskLocatorName)"
							}
						}
						Line 3 "List local hard drive in boot menu: " -nonewline
						if ($Device.localDiskEnabled -eq "1")
						{
							Write-Output -Message "Yes"
						} else {
							Write-Output -Message "No"
						}
						
						DeviceStatus $Device
					} else {
						Write-Output -Message "No Target Devices found. Device Collection is empty."
						Line 0 ''
						$obj1 = [PSCustomObject] @{
							CollectionName = $Collection.collectionName
						}
						$null = $Script:EmptyDeviceCollections.Add($obj1)
					}
				}
			}
		}
	}

	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function DeviceStatus
{
	Param($xDevice)

	if ($Null -eq $xDevice -or $xDevice.status -eq '' -or $xDevice.status -eq "0")
	{
		Write-Output -Message "Target device inactive"
	} else {
		Write-Output -Message "Target device active"
		Line 3 "IP Address: " $xDevice.ip
		Write-Output -Message "Server: $($xDevice.serverName)"
		Write-Output -Message "Server IP: $($xDevice.serverIpConnection)"
		Write-Output -Message "Server Port: $($xDevice.serverPortConnection)"
		Line 3 "vDisk: " $xDevice.diskLocatorName
		Line 3 "vDisk version: " $xDevice.diskVersion
		Line 3 "vDisk name: " $xDevice.diskFileName
		Line 3 "vDisk access: " -nonewline
		Switch ($xDevice.diskVersionAccess)
		{
			0 {Line 0 "Production"; Break}
			1 {Line 0 "Test"; Break}
			2 {Line 0 "Maintenance"; Break}
			3 {Line 0 "Personal vDisk"; Break}
			Default {Line 0 "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"; Break}
		}
		if ($PVSVersion -eq "7")
		{
			Write-Output -Message "Local write cache disk:$($xDevice.localWriteCacheDiskSize)GB"
			Line 3 "Boot mode:" -nonewline
			Switch($xDevice.bdmBoot)
			{
				0 {Line 0 "PXE boot"; Break}
				1 {Line 0 "BDM disk"; Break}
				Default {Line 0 "Boot mode could not be determined: $($xDevice.bdmBoot)"; Break}
			}
		}
		Switch($xDevice.licenseType)
		{
			0 {Line 3 "No License"; Break}
			1 {Line 3 "Desktop License"; Break}
			2 {Line 3 "Server License"; Break}
			5 {Line 3 "OEM SmartClient License"; Break}
			6 {Line 3 "XenApp License"; Break}
			7 {Line 3 "XenDesktop License"; Break}
			Default {Line 0 "Device license type could not be determined: $($xDevice.licenseType)"; Break}
		}
		
		Line 3 "Logging level: " -nonewline
		Switch ($xDevice.logLevel)
		{
			0   {Line 0 "Off"; Break}
			1   {Line 0 "Fatal"; Break}
			2   {Line 0 "Error"; Break}
			3   {Line 0 "Warning"; Break}
			4   {Line 0 "Info"; Break}
			5   {Line 0 "Debug"; Break}
			6   {Line 0 "Trace"; Break}
			Default {Line 0 "Logging level could not be determined: $($xDevice.logLevel)"; Break}
		}
	}
	Line 0 ''
}

Function GetBootstrapInfo
{
	Param([object]$server)

	Write-Verbose -Message ('{0}: Processing Bootstrap files' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Output -Message "Bootstrap settings"
	Write-Verbose -Message ('{0}: Processing Bootstrap files for Server $($server.servername)' -f (Get-Date -UFormat "%F %r (%Z)"))
	#first get all bootstrap files for the server
	$temp = $server.serverName
	$GetWhat = "ServerBootstrapNames"
	$GetParam = "serverName = $temp"
	$ErrorTxt = "Server Bootstrap Name information"
	$BootstrapNames = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	#Now that the list of bootstrap names has been gathered
	#We have the mandatory parameter to get the bootstrap info
	#there should be at least one bootstrap filename
	if ($Null -ne $Bootstrapnames)
	{
		#cannot use the BuildPVSObject Function here
		$serverbootstraps = @()
		ForEach($Bootstrapname in $Bootstrapnames)
		{
			#get serverbootstrap info
			$error.Clear()
			$tempserverbootstrap = Mcli-Get ServerBootstrap -p name="$($Bootstrapname.name)",servername="$($server.serverName)"
			if ($error.Count -eq 0)
			{
				$serverbootstrap = $Null
				ForEach($record in $tempserverbootstrap)
				{
					if ($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
					{
						if ($Null -ne $serverbootstrap)
						{
							$serverbootstraps +=  $serverbootstrap
						}
						$serverbootstrap = new-object System.Object
						#add the bootstrapname name value to the serverbootstrap object
						$property = "BootstrapName"
						$value = $Bootstrapname.name
						Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
					}
					$index = $record.IndexOf(':')
					if ($index -gt 0)
					{
						$property = $record.SubString(0, $index)
						$value = $record.SubString($index + 2)
						if ($property -ne "Executing")
						{
							Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
						}
					}
				}
				$serverbootstraps +=  $serverbootstrap
			} else {
				Write-Output -Message "Server Bootstrap information could not be retrieved"
				Line 2 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
			}
		}
		if ($Null -ne $ServerBootstraps)
		{
			Write-Verbose -Message ('{0}: Processing Bootstrap file $($ServerBootstrap.Bootstrapname)' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			ForEach($ServerBootstrap in $ServerBootstraps)
			{
				Write-Verbose -Message ('{0}: Gather Bootstrap info for Appendix D' -f (Get-Date -UFormat "%F %r (%Z)"))
				$obj1 = [PSCustomObject] @{
					ServerName 	  = $Server.serverName				
					BootstrapName = $ServerBootstrap.Bootstrapname				
					IP1        	  = $ServerBootstrap.bootserver1_Ip				
					IP2        	  = $ServerBootstrap.bootserver2_Ip				
					IP3        	  = $ServerBootstrap.bootserver3_Ip				
					IP4        	  = $ServerBootstrap.bootserver4_Ip				
				}
				$null = $Script:BootstrapItems.Add($obj1)

				Line 3 "Bootstrap file: " $ServerBootstrap.Bootstrapname
				if ($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address: " $ServerBootstrap.bootserver1_Ip
					Line 3 "Subnet Mask: " $ServerBootstrap.bootserver1_Netmask
					Line 3 "Gateway: " $ServerBootstrap.bootserver1_Gateway
					Line 3 "Port: " $ServerBootstrap.bootserver1_Port
				}
				if ($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address: " $ServerBootstrap.bootserver2_Ip
					Line 3 "Subnet Mask: " $ServerBootstrap.bootserver2_Netmask
					Line 3 "Gateway: " $ServerBootstrap.bootserver2_Gateway
					Line 3 "Port: " $ServerBootstrap.bootserver2_Port
				}
				if ($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address: " $ServerBootstrap.bootserver3_Ip
					Line 3 "Subnet Mask: " $ServerBootstrap.bootserver3_Netmask
					Line 3 "Gateway: " $ServerBootstrap.bootserver3_Gateway
					Line 3 "Port: " $ServerBootstrap.bootserver3_Port
				}
				if ($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address: " $ServerBootstrap.bootserver4_Ip
					Line 3 "Subnet Mask: " $ServerBootstrap.bootserver4_Netmask
					Line 3 "Gateway: " $ServerBootstrap.bootserver4_Gateway
					Line 3 "Port: " $ServerBootstrap.bootserver4_Port
				}
				Line 0 ''
				Write-Verbose -Message ('{0}: Processing Options Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				Line 3 "Verbose mode: " -nonewline
				if ($ServerBootstrap.verboseMode -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				Line 3 "Interrupt safe mode: " -nonewline
				if ($ServerBootstrap.interruptSafeMode -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				Line 3 "Advanced Memory Support: " -nonewline
				if ($ServerBootstrap.paeMode -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				Line 3 "Network recovery method: " -nonewline
				if ($ServerBootstrap.bootFromHdOnFail -eq "0")
				{
					Write-Output -Message "Restore network connection"
				} else {
					Write-Output -Message "Reboot to Hard Drive after $($ServerBootstrap.recoveryTime) seconds"
				}
				Line 3 "Login polling timeout: " -nonewline
				if ($ServerBootstrap.pollingTimeout -eq '')
				{
					Write-Output -Message "5000 (milliseconds)"
				} else {
					Write-Output -Message "$($ServerBootstrap.pollingTimeout) (milliseconds)"
				}
				Line 3 "Login general timeout: " -nonewline
				if ($ServerBootstrap.generalTimeout -eq '')
				{
					Write-Output -Message "5000 (milliseconds)"
				} else {
					Write-Output -Message "$($ServerBootstrap.generalTimeout) (milliseconds)"
				}
				Line 0 ''
			}
		}
	} else {
		Write-Output -Message "No Bootstrap names available"
	}
	Line 0 ''
}

Function GetPVSServiceInfo
{
	Param([string]$ComputerName)

	Write-Verbose -Message ('{0}: Processing PVS Services for Server $($server.servername)' -f (Get-Date -UFormat "%F %r (%Z)"))
	$Services = Get-WmiObject -ComputerName $ComputerName Win32_Service -EA 0 | `
	Where-Object {$_.DisplayName -like "Citrix PVS*"} | `
	Select-Object displayname, name, status, startmode, started, startname, state | `
	Sort-Object DisplayName
	
	if ($? -and $Null -ne $Services)
	{
		ForEach($Service in $Services)
		{
			$obj1 = [PSCustomObject] @{
				ServerName 	   = $ComputerName
				DisplayName	   = $Service.DisplayName
				Name  		   = $Service.Name
				Status 		   = $Service.Status
				StartMode  	   = $Service.StartMode
				Started		   = $Service.Started
				StartName  	   = $Service.StartName
				State  		   = $Service.State
				FailureAction1 = "Take no Action"
				FailureAction2 = "Take no Action"
				FailureAction3 = "Take no Action"
			}

			[array]$Actions = sc.exe \\$ComputerName qfailure $Service.Name
			
			if ($Actions.Length -gt 0)
			{
				if (($Actions -like "*RESTART -- Delay*") -or ($Actions -like "*RUN PROCESS -- Delay*") -or ($Actions -like "*REBOOT -- Delay*"))
				{
					$cnt = 0
					ForEach($Item in $Actions)
					{
						Switch ($Item)
						{
							{$Item -like "*RESTART -- Delay*"}		{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Restart the Service"; Break}
							{$Item -like "*RUN PROCESS -- Delay*"}	{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Run a Program"; Break}
							{$Item -like "*REBOOT -- Delay*"}		{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Restart the Computer"; Break}
						}
					}
				}
			}
			
			$null = $Script:PVSServiceItems.Add($obj1)
		}
	} else {
		Write-Output -Message "No PVS services found for $($ComputerName)"
	}
	Line 0 ''
}

Function GetPVSProcessInfo
{
	Param([string]$ComputerName)
	
	#Whether or not the Inventory executable is running (Inventory.exe)
	#Whether or not the Notifier executable is running (Notifier.exe)
	#Whether or not the MgmtDaemon executable is running (MgmtDaemon.exe)
	#Whether or not the StreamProcess executable is running (StreamProcess.exe)
	
	#All four of those run within the StreamService.exe process.

	Write-Verbose -Message ('{0}: Processing PVS Processes for Server $($server.servername)' -f (Get-Date -UFormat "%F %r (%Z)"))

	$InventoryProcess = Get-Process -Name 'Inventory' -ComputerName $ComputerName
	$NotifierProcess = Get-Process -Name 'Notifier' -ComputerName $ComputerName
	$MgmtDaemonProcess = Get-Process -Name 'MgmtDaemon' -ComputerName $ComputerName
	$StreamProcessProcess = Get-Process -Name 'StreamProcess' -ComputerName $ComputerName
	
	$tmp1 = "Inventory"
	$tmp2 = ''
	if ($InventoryProcess)
	{
		$tmp2 = "Running"
	} else {
		$tmp2 = "Not Running"
	}
	$obj1 = [PSCustomObject] @{
		ProcessName	= $tmp1
		ServerName 	= $ComputerName	
		Status  	= $tmp2
	}
	$null = $Script:PVSProcessItems.Add($obj1)
	
	$tmp1 = "Notifier"
	$tmp2 = ''
	if ($NotifierProcess)
	{
		$tmp2 = "Running"
	} else {
		$tmp2 = "Not Running"
	}
	$obj1 = [PSCustomObject] @{
		ProcessName	= $tmp1
		ServerName 	= $ComputerName	
		Status  	= $tmp2
	}
	$null = $Script:PVSProcessItems.Add($obj1)
	
	$tmp1 = "MgmtDaemon"
	$tmp2 = ''
	if ($MgmtDaemonProcess)
	{
		$tmp2 = "Running"
	} else {
		$tmp2 = "Not Running"
	}
	$obj1 = [PSCustomObject] @{
		ProcessName	= $tmp1
		ServerName 	= $ComputerName	
		Status  	= $tmp2
	}
	$null = $Script:PVSProcessItems.Add($obj1)
	
	$tmp1 = "StreamProcess"
	$tmp2 = ''
	if ($StreamProcessProcess)
	{
		$tmp2 = "Running"
	} else {
		$tmp2 = "Not Running"
	}
	$obj1 = [PSCustomObject] @{
		ProcessName	= $tmp1
		ServerName 	= $ComputerName	
		Status  	= $tmp2
	}
	$null = $Script:PVSProcessItems.Add($obj1)
}

Function GetBadStreamingIPAddresses
{
	Param([string]$ComputerName)
	#function updated by Andrew Williamson @ Fujitsu Services to handle servers with multiple NICs
	#further optimization by Michael B. Smith

	#loop through the configured streaming ip address and compare to the physical configured ip addresses
	#if a streaming ip address is not in the list of physical ip addresses, it is a bad streaming ip address
	ForEach ($Stream in ($Script:StreamingIPAddresses | Where-Object {$_.Servername -eq $ComputerName})) {
		$exists = $false
		:outerLoop ForEach ($ServerNIC in $Script:NICIPAddresses.Item($ComputerName)) 
		{
			ForEach ($IP in $ServerNIC) {
				# there could be more than one IP
				If ($Stream.IPAddress -eq $IP) 
				{
					$Exists = $true
					break :outerLoop
				}
			}
		}
		if (!$exists) 
		{
			$obj1 = [PSCustomObject] @{
				ServerName = $ComputerName			
				IPAddress  = $Stream.IPAddress			
			}
			$null = $Script:BadIPs.Add($obj1)
		}
	}
}

Function ProcessvDisksinFarm
{
	#process all vDisks in site
	Write-Verbose -Message ('{0}: Processing all vDisks in site' -f (Get-Date -UFormat "%F %r (%Z)"))
	[int]$NumberofvDisks = 0
	$GetWhat = "DiskInfo"
	$GetParam = ''
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	Write-Output -Message "vDisks in Farm"
	if ($Null -ne $Disks)
	{
		ForEach($Disk in $Disks)
		{
			Write-Verbose -Message ('{0}: Processing vDisk $($Disk.diskLocatorName)' -f (Get-Date -UFormat "%F %r (%Z)"))
			Line 1 $Disk.diskLocatorName
			if ($Script:PVSVersion -eq "5")
			{
				#PVS 5.x
				Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				Line 2 "Store: " $Disk.storeName
				Line 2 "Site: " $Disk.siteName
				Line 2 "Filename: " $Disk.diskLocatorName
				Line 2 "Size: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				Write-Output -Message " MB"
				if (![String]::IsNullOrEmpty($Disk.serverName))
				{
					Line 2 "Use this server to provide the vDisk: " $Disk.serverName
				} else {
					Line 2 "Subnet Affinity: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {Line 0 "None"; Break}
						1 {Line 0 "Best Effort"; Break}
						2 {Line 0 "Fixed"; Break}
						Default {Line 2 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"; Break}
					}
					Line 2 "Rebalance Enabled: " -nonewline
					if ($Disk.rebalanceEnabled -eq "1")
					{
						Write-Output -Message "Yes"
						Write-Output -Message "Trigger Percent: $($Disk.rebalanceTriggerPercent)"
					} else {
						Write-Output -Message "No"
					}
				}
				Line 2 "Allow use of this vDisk: " -nonewline
				if ($Disk.enabled -eq "1")
				{
					Write-Output -Message "Yes"
					if ($Disk.deviceCount -gt 0)
					{
						$NumberofvDisks++
					}
				} else {
					Write-Output -Message "No"
				}
				Line 2 "Access mode: " -nonewline
				if ($Disk.writeCacheType -eq "0")
				{
					Write-Output -Message "Private Image (single device, read/write access)"
				}
				elseif ($Disk.writeCacheType -eq "7")
				{
					Write-Output -Message "Difference Disk Image"
				} else {
					Write-Output -Message "Standard Image (multi-device, read-only access)"
					Line 2 "Cache type: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {Line 0 "Private Image"; Break}
						1   {
								Write-Output -Message "Cache on server"
								
								$obj1 = [PSCustomObject] @{
									StoreName = $Disk.storeName								
									SiteName  = $Disk.siteName								
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:CacheOnServer.Add($obj1)
								Break
							}
						2   {Line 0 "Cache encrypted on server disk"; Break}
						3   {
							Write-Output -Message "Cache in device RAM"
							Line 2 "Cache Size: $($Disk.writeCacheSize) MBs"; Break
							}
						4   {Line 0 "Cache on device's HD"; Break}
						5   {Line 0 "Cache encrypted on device's hard disk"; Break}
						6   {Line 0 "RAM Disk"; Break}
						7   {Line 0 "Difference Disk"; Break}
						Default {Line 0 "Cache type could not be determined: $($Disk.writeCacheType)"; Break}
					}
				}
				if ($Disk.activationDateEnabled -eq "0")
				{
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					if ($Disk.autoUpdateEnabled -eq "1")
					{
						Write-Output -Message "Yes"
					} else {
						Write-Output -Message "No"
					}
					Write-Output -Message "Apply vDisk updates as soon as they are detected by the server"
				} else {
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					if ($Disk.autoUpdateEnabled -eq "1")
					{
						Write-Output -Message "Yes"
					} else {
						Write-Output -Message "No"
					}
					Write-Output -Message "Schedule the next vDisk update to occur on: $($Disk.activeDate)"
				}
				Line 2 "Microsoft license type: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {Line 0 "None"; Break}
					1 {Line 0 "Multiple Activation Key (MAK)"; Break}
					2 {Line 0 "Key Management Service (KMS)"; Break}
					Default {Line 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"; Break}
				}
				#options tab
				Write-Verbose -Message ('{0}: Processing Options Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				Line 2 "High availability (HA): " -nonewline
				if ($Disk.haEnabled -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				Line 2 "AD machine account password management: " -nonewline
				if ($Disk.adPasswordEnabled -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				
				Line 2 "Printer management: " -nonewline
				if ($Disk.printerManagementEnabled -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
			} else {
				#PVS 6.x or 7.x
				Write-Verbose -Message ('{0}: Processing vDisk Properties' -f (Get-Date -UFormat "%F %r (%Z)"))
				Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				Line 2 "Site: " $Disk.siteName
				Line 2 "Store: " $Disk.storeName
				Line 2 "Filename: " $Disk.diskLocatorName
				Line 2 "Size: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				Write-Output -Message " MB"
				Line 2 "VHD block size: " $Disk.vhdBlockSize -nonewline
				Write-Output -Message " KB"
				Line 2 "Access mode: " -nonewline
				if ($Disk.writeCacheType -eq "0")
				{
					Write-Output -Message "Private Image (single device, read/write access)"
				} else {
					Write-Output -Message "Standard Image (multi-device, read-only access)"
					Line 2 "Cache type: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {Line 0 "Private Image"; Break}
						1   {
								Write-Output -Message "Cache on server"
								
								$obj1 = [PSCustomObject] @{
									StoreName = $Disk.storeName								
									SiteName  = $Disk.siteName								
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:CacheOnServer.Add($obj1)
								Break
							}
						3   {
							Write-Output -Message "Cache in device RAM"
							Line 2 "Cache Size: $($Disk.writeCacheSize) MBs"; Break
							}
						4   {Line 0 "Cache on device's hard disk"; Break}
						6   {Line 0 "RAM Disk"; Break}
						7   {Line 0 "Difference Disk"; Break}
						8  	{Line 0 "Cache on device hard drive persisted (NT 6.1 and later)"; Break}
						9  	{Line 0 "Cache in device RAM with overflow on hard disk"; Break}
						10 	{Line 0 "Private Image with Asynchronous IO"; Break} #added 1811
						11 	{Line 0 "Cache on server, persistent with Asynchronous IO"; Break} #added 1811
						12 	{Line 0 "Cache in device RAM with overflow on hard disk with Asynchronous IO"; Break} #added 1811
						Default {Line 0 "Cache type could not be determined: $($Disk.writeCacheType)"; Break}
					}
				}
				if (![String]::IsNullOrEmpty($Disk.menuText))
				{
					Line 2 "BIOS boot menu text: " $Disk.menuText
				}
				Line 2 "Enable AD machine acct pwd mgmt: " -nonewline
				if ($Disk.adPasswordEnabled -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				
				Line 2 "Enable printer management: " -nonewline
				if ($Disk.printerManagementEnabled -eq "1")
				{
					Write-Output -Message "Yes"
				} else {
					Write-Output -Message "No"
				}
				Line 2 "Enable streaming of this vDisk: " -nonewline
				if ($Disk.Enabled -eq "1")
				{
					Write-Output -Message "Yes"
					if ($Disk.deviceCount -gt 0)
					{
						$NumberofvDisks++
					}
				} else {
					Write-Output -Message "No"
				}
				Line 2 "Microsoft license type: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {Line 0 "None"; Break}
					1 {Line 0 "Multiple Activation Key (MAK)"; Break}
					2 {Line 0 "Key Management Service (KMS)"; Break}
					Default {Line 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"; Break}
				}
				if ($Script:PVSFullVersion -ge "7.22")
				{
					if ($Disk.AccelerateOfficeActivation)
					{
						Write-Output -Message "Accelerate Office Activation: Yes"
					} else {
						Write-Output -Message "Accelerate Office Activation: No"
					}
				}

				Write-Verbose -Message ('{0}: Processing Auto Update Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
				if ($Disk.activationDateEnabled -eq "0")
				{
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					if ($Disk.autoUpdateEnabled -eq "1")
					{
						Write-Output -Message "Yes"
					} else {
						Write-Output -Message "No"
					}
					Write-Output -Message "Apply vDisk updates as soon as they are detected by the server"
				} else {
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					if ($Disk.autoUpdateEnabled -eq "1")
					{
						Write-Output -Message "Yes"
					} else {
						Write-Output -Message "No"
					}
					Write-Output -Message "Schedule the next vDisk update to occur on: $($Disk.activeDate)"
				}
				#process Versions menu
				#get versions info
				#thanks to the PVS Product team for their help in understanding the Versions information
				Write-Verbose -Message ('{0}: Processing vDisk Versions' -f (Get-Date -UFormat "%F %r (%Z)"))
				$error.Clear()
				$MCLIGetResult = Mcli-Get DiskVersion -p diskLocatorName="$($Disk.diskLocatorName)",storeName="$($disk.storeName)",siteName="$($disk.siteName)"
				if ($error.Count -eq 0)
				{
					#build versions object
					$PluralObject = @()
					$SingleObject = $Null
					ForEach($record in $MCLIGetResult)
					{
						if ($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
						{
							if ($Null -ne $SingleObject)
							{
								$PluralObject += $SingleObject
							}
							$SingleObject = new-object System.Object
						}

						$index = $record.IndexOf(':')
						if ($index -gt 0)
						{
							$property = $record.SubString(0, $index)
							$value    = $record.SubString($index + 2)
							if ($property -ne "Executing")
							{
								Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
							}
						}
					}
					$PluralObject += $SingleObject
					$DiskVersions = $PluralObject
					
					if ($Null -ne $DiskVersions)
					{
						#get the current booting version
						#by default, the $DiskVersions object is in version number order lowest to highest
						#the initial or base version is 0 and always exists
						[string]$BootingVersion = "0"
						[bool]$BootOverride = $False
						ForEach($DiskVersion in $DiskVersions)
						{
							if ($DiskVersion.access -eq "3")
							{
								#override i.e. manually selected boot version
								$BootingVersion = $DiskVersion.version
								$BootOverride = $True
								Break
							} elseif ($DiskVersion.access -eq "0" -and $DiskVersion.IsPending -eq "0" ) {
								$BootingVersion = $DiskVersion.version
								$BootOverride = $False
							}
						}
						
						Line 2 "Boot production devices from version: " -NoNewLine
						if ($BootOverride)
						{
							Line 0 $BootingVersion
						} else {
							Write-Output -Message "Newest released"
						}
						Line 0 ''
						
						$VersionFlag = $False
						ForEach($DiskVersion in $DiskVersions)
						{
							Write-Verbose -Message ('{0}: Processing vDisk Version $($DiskVersion.version)' -f (Get-Date -UFormat "%F %r (%Z)"))
							Line 2 "Version: " -NoNewLine
							if ($DiskVersion.version -eq $BootingVersion)
							{
								Write-Output -Message "$($DiskVersion.version) (Current booting version)"
							} else {
								Line 0 $DiskVersion.version
							}
							if ($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
							{
								$VersionFlag = $True
								Write-Output -Message "Version of vDisk is $($DiskVersion.version) which is greater than the limit of $($Script:farm.maxVersions). Consider merging."
								
								$obj1 = [PSCustomObject] @{
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:VersionsToMerge.Add($obj1)
							}
							Line 2 "Created: " $DiskVersion.createDate
							if (![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								Line 2 "Released: " $DiskVersion.scheduledDate
							}
							Line 2 "Devices: " $DiskVersion.deviceCount
							Line 2 "Access: " -NoNewLine
							Switch ($DiskVersion.access)
							{
								"0" {Line 0 "Production"; Break}
								"1" {Line 0 "Maintenance"; Break}
								"2" {Line 0 "Maintenance Highest Version"; Break}
								"3" {Line 0 "Override"; Break}
								"4" {Line 0 "Merge"; Break}
								"5" {Line 0 "Merge Maintenance"; Break}
								"6" {Line 0 "Merge Test"; Break}
								"7" {Line 0 "Test"; Break}
								Default {Line 0 "Access could not be determined: $($DiskVersion.access)"; Break}
							}
							Line 2 "Type: " -NoNewLine
							Switch ($DiskVersion.type)
							{
								"0" {Line 0 "Base"; Break}
								"1" {Line 0 "Manual"; Break}
								"2" {Line 0 "Automatic"; Break}
								"3" {Line 0 "Merge"; Break}
								"4" {Line 0 "Merge Base"; Break}
								Default {Line 0 "Type could not be determined: $($DiskVersion.type)"; Break}
							}
							if (![String]::IsNullOrEmpty($DiskVersion.description))
							{
								Line 2 "Properties: " $DiskVersion.description
							}
							Line 2 "Can Delete: "  -NoNewLine
							Switch ($DiskVersion.canDelete)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Merge: "  -NoNewLine
							Switch ($DiskVersion.canMerge)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Merge Base: "  -NoNewLine
							Switch ($DiskVersion.canMergeBase)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Promote: "  -NoNewLine
							Switch ($DiskVersion.canPromote)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Revert back to Test: "  -NoNewLine
							Switch ($DiskVersion.canRevertTest)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Revert back to Maintenance: "  -NoNewLine
							Switch ($DiskVersion.canRevertMaintenance)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Set Scheduled Date: "  -NoNewLine
							Switch ($DiskVersion.canSetScheduledDate)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Override: "  -NoNewLine
							Switch ($DiskVersion.canOverride)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Is Pending: "  -NoNewLine
							Switch ($DiskVersion.isPending)
							{
								0 {Line 0 "No, version Scheduled Date has occurred"; Break}
								1 {Line 0 "Yes, version Scheduled Date has not occurred"; Break}
							}
							Line 2 "Replication Status: " -NoNewLine
							Switch ($DiskVersion.goodInventoryStatus)
							{
								0 {Line 0 "Not available on all servers"; Break}
								1 {Line 0 "Available on all servers"; Break}
								Default {Line 0 "Replication status could not be determined: $($DiskVersion.goodInventoryStatus)"; Break}
							}
							Line 2 "Disk Filename: " $DiskVersion.diskFileName
							Line 0 ''
						}
					}
				} else {
					Write-Output -Message "Disk Version information could not be retrieved"
					Line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
				
				#process vDisk Load Balancing Menu
				Write-Verbose -Message ('{0}: Processing vDisk Load Balancing Menu' -f (Get-Date -UFormat "%F %r (%Z)"))
				if (![String]::IsNullOrEmpty($Disk.serverName))
				{
					Line 2 "Use this server to provide the vDisk: " $Disk.serverName
				} else {
					Line 2 "Subnet Affinity: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {Line 0 "None"; Break}
						1 {Line 0 "Best Effort"; Break}
						2 {Line 0 "Fixed"; Break}
						Default {Line 0 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"; Break}
					}
					Line 2 "Rebalance Enabled: " -nonewline
					if ($Disk.rebalanceEnabled -eq "1")
					{
						Write-Output -Message "Yes"
						Write-Output -Message "Trigger Percent: $($Disk.rebalanceTriggerPercent)"
					} else {
						Write-Output -Message "No"
					}
				}
			}
			Line 0 ''
		}
	}

	Write-Output -Message "Number of vDisks that are Enabled and have active connections: " $NumberofvDisks
	Line 0 ''
	# http://blogs.citrix.com/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/
	[decimal]$RecRAM = ((2 + ($NumberofvDisks * 2)) * 1.15)
	$RecRAM = "{0:N0}" -f $RecRAM
	Write-Output -Message "Recommended RAM for each PVS Server using XenDesktop vDisks: $($RecRAM)GB"
	[decimal]$RecRAM = ((2 + ($NumberofvDisks * 4)) * 1.15)
	$RecRAM = "{0:N0}" -f $RecRAM
	Write-Output -Message "Recommended RAM for each PVS Server using XenApp vDisks: $($RecRAM)GB"
	Line 0 ''
	Write-Output -Message "This script is not able to tell if a vDisk is running XenDesktop or XenApp."
	Write-Output -Message "The RAM calculation is done based on both scenarios. The original formula is:"
	Write-Output -Message "2GB + (#XA_vDisk * 4GB) + (#XD_vDisk * 2GB) + 15% (Buffer)"
	Line 1 'PVS Internals 2 - How to properly size your memory by Martin Zugec'
	Line 1 'https://www.citrix.com/blogs/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/'
	Line 0 ''
}

Function GetMicrosoftHotfixes 
{
	Param([string]$ComputerName)
	
	#added V1.16 get installed Microsoft Hotfixes and Updates
	Write-Verbose -Message ('{0}: Retrieving Microsoft hotfixes and updates' -f (Get-Date -UFormat "%F %r (%Z)"))
	[bool]$GotMSHotfixes = $True
	
	Try
	{
		$results = Get-HotFix -ComputerName $ComputerName | Select-Object CSName,Caption,Description,HotFixID,InstalledBy,InstalledOn
		$MSInstalledHotfixes = $results | Sort-Object HotFixID
		$results = $Null
	}
	
	Catch
	{
		$GotMSHotfixes = $False
	}

	Write-Verbose -Message ('{0}: Output Microsoft hotfixes and updates' -f (Get-Date -UFormat "%F %r (%Z)"))
	if ($GotMSHotfixes -eq $False)
	{
		Write-Output -Message "No installed Microsoft hotfixes or updates were found"
		Line 0 ''
	} else {
		ForEach($Hotfix in $MSInstalledHotfixes)
		{
			$obj1 = [PSCustomObject] @{
				HotFixID	= $Hotfix.HotFixID			
				ServerName	= $Hotfix.CSName			
				Caption		= $Hotfix.Caption			
				Description	= $Hotfix.Description			
				InstalledBy	= $Hotfix.InstalledBy			
				InstalledOn	= $Hotfix.InstalledOn			
			}
			$null = $Script:MSHotfixes.Add($obj1)
		}
	}
}

Function GetInstalledRolesAndFeatures
{
	Param([string]$ComputerName)
	
	#added V1.16 get Windows installed Roles and Features
	Write-Verbose -Message ('{0}: Retrieving Windows installed Roles and Features' -f (Get-Date -UFormat "%F %r (%Z)"))
	[bool]$GotWinComponents = $True
	
	$results = Get-WindowsFeature -ComputerName $ComputerName -EA 0 4> $Null
	
	if (!$?)
	{
		$GotWinComponents = $False
	}
	
	$WinComponents = $results | Where-Object Installed | Select-Object DisplayName,Name,FeatureType | Sort-Object DisplayName 
	
	Write-Verbose -Message ('{0}: Output Windows installed Roles and Features' -f (Get-Date -UFormat "%F %r (%Z)"))
	if ($GotWinComponents -eq $False)
	{
		Write-Output -Message "No Windows installed Roles and Features were found"
		Line 0 ''
	} else {
		ForEach($Component in $WinComponents)
		{
			$obj1 = [PSCustomObject] @{
				DisplayName	= $Component.DisplayName			
				Name		= $Component.Name			
				ServerName	= $ComputerName			
				FeatureType	= $Component.FeatureType			
			}
			$null = $Script:WinInstalledComponents.Add($obj1)
		}
	}
}

Function ProcessStores
{
	#process the stores now
	Write-Verbose -Message ('{0}: Processing Stores' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Output -Message "Stores Properties"
	$GetWhat = "Store"
	$GetParam = ''
	$ErrorTxt = "Farm Store information"
	$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	if ($Null -ne $Stores)
	{
		ForEach($Store in $Stores)
		{
			Write-Verbose -Message ('{0}: Processing Store $($Store.StoreName)' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Verbose -Message ('{0}: Processing General Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Output -Message "Name: " $Store.StoreName
			
			Write-Verbose -Message ('{0}: Processing Servers Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Output -Message "Servers"
			#find the servers (and the site) that serve this store
			$GetWhat = "Server"
			$GetParam = ''
			$ErrorTxt = "Server information"
			$Servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			$StoreServers = @()
			if ($Null -ne $Servers)
			{
				ForEach($Server in $Servers)
				{
					Write-Verbose -Message ('{0}: Processing Server $($Server.serverName)' -f (Get-Date -UFormat "%F %r (%Z)"))
					$Temp = $Server.serverName
					$GetWhat = "ServerStore"
					$GetParam = "serverName = $Temp"
					$ErrorTxt = "Server Store information"
					$ServerStore = BuildPVSObject $GetWhat $GetParam $ErrorTxt
                    $Providers = $ServerStore | Where-Object {$_.StoreName -eq $Store.Storename}
                    if ($Providers)
					{
                       ForEach ($Provider in $Providers)
					   {
                          $StoreServers += $Provider.ServerName
                       }
                    }
				}	
			}
			Write-Output -Message "Servers that provide this store:"
			ForEach($StoreServer in $StoreServers)
			{
				Line 3 $StoreServer
			}

			Write-Verbose -Message ('{0}: Processing Paths Tab' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Output -Message "Paths"

			if (Test-Path $Store.path -EA 0)
			{
				Write-Output -Message "Default store path: $($Store.path)"
			} else {
				Write-Output -Message "Default store path: $($Store.path) (Invalid path)"
			}
			
			if (![String]::IsNullOrEmpty($Store.cachePath))
			{
				Write-Output -Message "Default write-cache paths: "
				$WCPaths = @($Store.cachePath.Split(","))
				ForEach($WCPath in $WCPaths)
				{
					if (Test-Path $WCPath -EA 0)
					{
						Line 3 $WCPath
					} else {
						Write-Output -Message "$($WCPath) (Invalid path)"
					}
					#Line 3 $WCPath
				}
			}
			Line 0 ''
		}
	} else {
		Write-Output -Message "There are no Stores configured"
	}
	Line 0 ''
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixA
{
	Write-Verbose -Message ('{0}: Create Appendix A Advanced Server Items (Server/Network)' -f (Get-Date -UFormat "%F %r (%Z)"))
	#sort the array by servername
	$Script:AdvancedItems1 = $Script:AdvancedItems1 | Sort-Object ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixA_AdvancedServerItems1.csv"
		$Script:AdvancedItems1 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Write-Output -Message "Appendix A - Advanced Server Items (Server/Network)"
	Line 0 ''
	Write-Output -Message "Server Name      Threads  Buffers  Server   Local       Remote      Ethernet  IO     Enable      "
	Write-Output -Message "                 per      per      Cache    Concurrent  Concurrent  MTU       Burst  Non-blocking"
	Write-Output -Message "                 Port     Thread   Timeout  IO Limit    IO Limit              Size   IO          "
	Write-Output -Message "================================================================================================="

	ForEach($Item in $Script:AdvancedItems1)
	{
		Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-11} {5,-11} {6,-9} {7,-6} {8,-8}" -f `
		$Item.serverName, $Item.threadsPerPort, $Item.buffersPerThread, $Item.serverCacheTimeout, `
		$Item.localConcurrentIoLimit, $Item.remoteConcurrentIoLimit, $Item.maxTransmissionUnits, $Item.ioBurstSize, `
		$Item.nonBlockingIoEnabled )
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix A - Advanced Server Items (Server/Network)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixB
{
	Write-Verbose -Message ('{0}: Create Appendix B Advanced Server Items (Pacing/Device)' -f (Get-Date -UFormat "%F %r (%Z)"))
	#sort the array by servername
	$Script:AdvancedItems2 = $Script:AdvancedItems2 | Sort-Object ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixB_AdvancedServerItems2.csv"
		$Script:AdvancedItems2 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Write-Output -Message "Appendix B - Advanced Server Items (Pacing/Device)"
	Line 0 ''
	Write-Output -Message "Server Name      Boot     Maximum  Maximum  vDisk     License"
	Write-Output -Message "                 Pause    Boot     Devices  Creation  Timeout"
	Write-Output -Message "                 Seconds  Time     Booting  Pacing           "
	Write-Output -Message "============================================================="
	###### "123451234512345  9999999  9999999  9999999  99999999  9999999

	ForEach($Item in $Script:AdvancedItems2)
	{
		Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-9} {5,-8}" -f `
		$Item.serverName, $Item.bootPauseSeconds, $Item.maxBootSeconds, $Item.maxBootDevicesAllowed, `
		$Item.vDiskCreatePacing, $Item.licenseTimeout )
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix B - Advanced Server Items (Pacing/Device)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixC
{
	Write-Verbose -Message ('{0}: Create Appendix C Config Wizard Items' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by servername
	$Script:ConfigWizItems = $Script:ConfigWizItems | Sort-Object ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixC_ConfigWizardItems.csv"
		$Script:ConfigWizItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix C - Configuration Wizard Settings"
	Line 0 ''
	Write-Output -Message "Server Name      DHCP        PXE       TFTP    User                                               " 
	Write-Output -Message "                 Services    Services  Option  Account                                            "
	Write-Output -Message "================================================================================================"

	if ($Script:ConfigWizItems)
	{
		ForEach($Item in $Script:ConfigWizItems)
		{
			Line 1 ( "{0,-16} {1,-11} {2,-9} {3,-7} {4,-50}" -f `
			$Item.serverName, $Item.DHCPServicesValue, $Item.PXEServicesValue, $Item.TFTPOptionValue, `
			$Item.UserAccount )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''
	Write-Verbose -Message ('{0}: Finished Creating Appendix C - Config Wizard Items' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixD
{
	Write-Verbose -Message ('{0}: Create Appendix D Server Bootstrap Items' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by bootstrapname and servername
	$Script:BootstrapItems = $Script:BootstrapItems | Sort-Object BootstrapName, ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixD_ServerBootstrapItems.csv"
		$Script:BootstrapItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix D - Server Bootstrap Items"
	Line 0 ''
	Write-Output -Message "Bootstrap Name   Server Name      IP1              IP2              IP3              IP4" 
	Write-Output -Message "===================================================================================================="
    ########123456789012345  XXXXXXXXXXXXXXXX 123.123.123.123  123.123.123.123  123.123.123.123  123.123.123.123
	if ($Script:BootstrapItems)
	{
		ForEach($Item in $Script:BootstrapItems)
		{
			Line 1 ( "{0,-16} {1,-16} {2,-16} {3,-16} {4,-16} {5,-16}" -f `
			$Item.BootstrapName, $Item.serverName, $Item.IP1, $Item.IP2, $Item.IP3, $Item.IP4 )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix D - Server Bootstrap Items' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixE
{
	Write-Verbose -Message ('{0}: Create Appendix E DisableTaskOffload Setting' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by bootstrapname and servername
	$Script:TaskOffloadItems = $Script:TaskOffloadItems | Sort-Object ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixE_DisableTaskOffloadSetting.csv"
		$Script:TaskOffloadItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix E - DisableTaskOffload Settings"
	Line 0 ''
	Write-Output -Message "Best Practices for Configuring Provisioning Services Server on a Network"
	Write-Output -Message "http://support.citrix.com/article/CTX117374"
	Write-Output -Message "This setting is not needed if you are running PVS 6.0 or later"
	Line 0 ''
	Write-Output -Message "Server Name      DisableTaskOffload Setting" 
	Write-Output -Message "==========================================="
	if ($Script:TaskOffloadItems)
	{
		ForEach($Item in $Script:TaskOffloadItems)
		{
			Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.TaskOffloadValue )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix E - DisableTaskOffload Setting' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixF
{
	Write-Verbose -Message ('{0}: Create Appendix F PVS Services' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by displayname and servername
	$Script:PVSServiceItems = $Script:PVSServiceItems | Sort-Object DisplayName, ServerName
	
	if ($CSV)
	{
		#AppendixF and AppendixF2 items are contained in the same array
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixF_PVSServices.csv"
		$Script:PVSServiceItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix F - Server PVS Service Items"
	Line 0 ''
	Write-Output -Message "Display Name                      Server Name      Service Name  Status Startup Type Started State   Log on as" 
	Write-Output -Message "========================================================================================================================================"
    ########123456789012345678901234567890123 123456789012345  1234567890123 123456 123456789012 1234567 
	#displayname, servername, name, status, startmode, started, startname, state 
	if ($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-6} {4,-12} {5,-7} {6,-7} {7,-35}" -f `
			$Item.DisplayName, $Item.serverName, $Item.Name, $Item.Status, $Item.StartMode, `
			$Item.Started, $Item.State, $Item.StartName )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix F - PVS Services' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixF2
{
	Write-Verbose -Message ('{0}: Create Appendix F2 PVS Services Failure Actions' -f (Get-Date -UFormat "%F %r (%Z)"))
	#array is already sorted in Function OutputAppendixF
	
	Write-Output -Message "Appendix F2 - Server PVS Service Items Failure Actions"
	Line 0 ''
	Write-Output -Message "Display Name                      Server Name      Service Name  Failure Action 1     Failure Action 2     Failure Action 3    " 
	Write-Output -Message "==============================================================================================================================="
	if ($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-20} {4,-20} {5,-20}" -f `
			$Item.DisplayName, $Item.serverName, $Item.Name, $Item.FailureAction1, $Item.FailureAction2, $Item.FailureAction3 )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix F2 - PVS Services Failure Actions' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixG
{
	Write-Verbose -Message ('{0}: Create Appendix G vDisks to Merge' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array
	$Script:VersionsToMerge = $Script:VersionsToMerge | Sort-Object
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixG_vDiskstoMerge.csv"
		$Script:VersionsToMerge | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix G - vDisks to Consider Merging"
	Line 0 ''
	Line 1 "vDisk Name" 
	Write-Output -Message "========================================"
	if ($Script:VersionsToMerge)
	{
		ForEach($Item in $Script:VersionsToMerge)
		{
			Line 1 ( "{0,-40}" -f $Item )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix G - vDisks to Merge' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixH
{
	Write-Verbose -Message ('{0}: Create Appendix H Empty Device Collections' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array
	$Script:EmptyDeviceCollections = $Script:EmptyDeviceCollections | Sort-Object CollectionName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixH_EmptyDeviceCollections.csv"
		$Script:EmptyDeviceCollections | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix H - Empty Device Collections"
	Line 0 ''
	Line 1 "Device Collection Name" 
	Write-Output -Message "=================================================="
	if ($Script:EmptyDeviceCollections)
	{
		ForEach($Item in $Script:EmptyDeviceCollections)
		{
			Line 1 ( "{0,-50}" -f $Item.CollectionName )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix G - Empty Device Collections' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function ProcessvDisksWithNoAssociation
{
	Write-Verbose -Message ('{0}: Finding vDisks with no Target Device Associations' -f (Get-Date -UFormat "%F %r (%Z)"))
	$UnassociatedvDisks = New-Object System.Collections.ArrayList
	$GetWhat = "diskLocator"
	$GetParam = ''
	$ErrorTxt = "Disk Locator information"
	$DiskLocators = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	if ($Null -eq $DiskLocators)
	{
		Write-Host -foregroundcolor Red -backgroundcolor Black "VERBOSE: $(Get-Date): No DiskLocators Found"
		OutputAppendixI $Null
	} else {
		ForEach($DiskLocator in $DiskLocators)
		{
			#get the diskLocatorId
			$DiskLocatorId = $DiskLocator.diskLocatorId
			
			#now pass the disklocatorid to get device
			#if nothing found, the vDisk is unassociated
			$temp = $DiskLocatorId
			$GetWhat = "device"
			$GetParam = "diskLocatorId = $temp"
			$ErrorTxt = "Device for DiskLocatorId $DiskLocatorId information"
			$Results = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			if ($Null -ne $Results)
			{
				#device found, vDisk is associated
			} else {
				#no device found that uses this vDisk
				$obj1 = [PSCustomObject] @{
					vDiskName = $DiskLocator.diskLocatorName				
				}
				$null = $UnassociatedvDisks.Add($obj1)
			}
		}
		
		if ($UnassociatedvDisks.Count -gt 0)
		{
			Write-Verbose -Message ('{0}: Found $($UnassociatedvDisks.Count) vDisks with no Target Device Associations' -f (Get-Date -UFormat "%F %r (%Z)"))
			OutputAppendixI $UnassociatedvDisks
		} else {
			Write-Verbose -Message ('{0}: All vDisks have Target Device Associations' -f (Get-Date -UFormat "%F %r (%Z)"))
			Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
			OutputAppendixI $Null
		}
	}
}

Function OutputAppendixI
{
	Param([array]$vDisks)

	Write-Verbose -Message ('{0}: Create Appendix I Unassociated vDisks' -f (Get-Date -UFormat "%F %r (%Z)"))

	Write-Output -Message "Appendix I - vDisks with no Target Device Associations"
	Line 0 ''
	Line 1 "vDisk Name" 
	Write-Output -Message "========================================"
	
	if ($vDisks)
	{
		#sort the array
		$vDisks = $vDisks | Sort-Object
	
		if ($CSV)
		{
			$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixI_UnassociatedvDisks.csv"
			$vDisks | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
		}
	
		ForEach($Item in $vDisks)
		{
			Line 1 ( "{0,-40}" -f $Item.vDiskName )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix I - Unassociated vDisks' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixJ
{
	Write-Verbose -Message ('{0}: Create Appendix J Bad Streaming IP Addresses' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by bootstrapname and servername
	$Script:BadIPs = $Script:BadIPs | Sort-Object ServerName, IPAddress
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixJ_BadStreamingIPAddresses.csv"
		$Script:BadIPs | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix J - Bad Streaming IP Addresses"
	Write-Output -Message "Streaming IP addresses that do not exist on the server"
	Line 0 ''
	Line 1 "Server Name      Streaming IP Address" 
	Write-Output -Message "====================================="
	if ($Script:BadIPs) 
	{
		ForEach($Item in $Script:BadIPs)
		{
			Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.IPAddress )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix J Bad Streaming IP Addresses' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixK
{
	Write-Verbose -Message ('{0}: Create Appendix K Misc Registry Items' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by regkey, regvalue and servername
	$Script:MiscRegistryItems = $Script:MiscRegistryItems | Sort-Object RegKey, RegValue, ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixK_MiscRegistryItems.csv"
		$Script:MiscRegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix K - Misc Registry Items"
	Write-Output -Message "Miscellaneous Registry Items That May or May Not Exist on Servers"
	Write-Output -Message "These items may or may not be needed"
	Write-Output -Message "This Appendix is strictly for server comparison only"
	Line 0 ''
	Line 1 "Registry Key                                                                                    Registry Value                                     Data                                                                                       Server Name    " 
	Write-Output -Message "============================================================================================================================================================================================================================================================="
	#       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345S12345678901234567890123456789012345678901234567890S123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890S123456789012345
	
	$Save = ''
	$First = $True
	if ($Script:MiscRegistryItems)
	{
		ForEach($Item in $Script:MiscRegistryItems)
		{
			if (!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())")
			{
				Line 0 ''
			}

			Line 1 ( "{0,-95} {1,-50} {2,-90} {3,-15}" -f `
			$Item.RegKey, $Item.RegValue, $Item.Value, $Item.serverName )
			$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())"
			if ($First)
			{
				$First = $False
			}
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix K Misc Registry Items' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixL
{
	Write-Verbose -Message ('{0}: Create Appendix L vDisks Configured for Server-Side Caching' -f (Get-Date -UFormat "%F %r (%Z)"))
	#sort the array 
	$Script:CacheOnServer = $Script:CacheOnServer | Sort-Object StoreName,SiteName,vDiskName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixL_vDisksConfiguredforServerSideCaching.csv"
		$Script:CacheOnServer | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Write-Output -Message "Appendix L - vDisks Configured for Server Side-Caching"
	Line 0 ''

	if ($Script:CacheOnServer)
	{
		Write-Output -Message "Store Name                Site Name                 vDisk Name               "
		Write-Output -Message "============================================================================="
			   #1234567890123456789012345 1234567890123456789012345 1234567890123456789012345

		ForEach($Item in $Script:CacheOnServer)
		{
			Line 1 ( "{0,-25} {1,-25} {2,-25}" -f `
			$Item.StoreName, $Item.SiteName, $Item.vDiskName )
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''
	
	Write-Verbose -Message ('{0}: Finished Creating Appendix L vDisks Configured for Server-Side Caching' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixM
{
	#added in V1.16
	Write-Verbose -Message ('{0}: Create Appendix M Microsoft Hotfixes and Updates' -f (Get-Date -UFormat "%F %r (%Z)"))

	#sort the array by hotfixid and servername
	$Script:MSHotfixes = $Script:MSHotfixes | Sort-Object HotFixID, ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixM_MicrosoftHotfixesandUpdates.csv"
		$Script:MSHotfixes | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix M - Microsoft Hotfixes and Updates"
	Line 0 ''
	Write-Output -Message "Hotfix ID                 Server Name     Caption                                       Description          Installed By                        Installed On Date     "
	Write-Output -Message "======================================================================================================================================================================="
	#       1234567890123456789012345S123456789012345S123456789012345678901234567890123456789012345S12345678901234567890S12345678901234567890123456789012345S1234567890123456789012
	#                                                 http://support.microsoft.com/?kbid=2727528    Security Update      XXX-XX-XDDC01\xxxx.xxxxxx           00/00/0000 00:00:00 PM
	#		25                        15              45                                            20                   35                                  22
	
	$Save = ''
	$First = $True
	if ($Script:MSHotfixes)
	{
		ForEach($Item in $Script:MSHotfixes)
		{
			if (!$First -and $Save -ne "$($Item.HotFixID)")
			{
				Line 0 ''
			}

			Line 1 ( "{0,-25} {1,-15} {2,-45} {3,-20} {4,-35} {5,-22}" -f `
			$Item.HotFixID, $Item.ServerName, $Item.Caption, $Item.Description, $Item.InstalledBy, $Item.InstalledOn)
			$Save = "$($Item.HotFixID)"
			if ($First)
			{
				$First = $False
			}
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix M Microsoft Hotfixes and Updates' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixN
{
	#added in V1.16
	Write-Verbose -Message ('{0}: Create Appendix N Windows Installed Components' -f (Get-Date -UFormat "%F %r (%Z)"))

	$Script:WinInstalledComponents = $Script:WinInstalledComponents | Sort-Object DisplayName, Name, DDCName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixN_InstalledRolesandFeatures.csv"
		$Script:WinInstalledComponents | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix N - Windows Installed Components"
	Line 0 ''
	Write-Output -Message "Display Name                                       Name                          Server Name      Feature Type   "
	Write-Output -Message "================================================================================================================="
	#       12345678901234567890123456789012345678901234567890S123456789012345678901234567890123456789012345SS123456789012345
	#       Graphical Management Tools and Infrastructure      NET-Framework-45-Features     XXXXXXXXXXXXXXX  Role Service
	#       50                                                 30                            15               15
	$Save = ''
	$First = $True
	if ($Script:WinInstalledComponents)
	{
		ForEach($Item in $Script:WinInstalledComponents)
		{
			if (!$First -and $Save -ne "$($Item.DisplayName)$($Item.Name)")
			{
				Line 0 ''
			}

			Line 1 ( "{0,-50} {1,-30} {2,-15} {3,-15}" -f `
			$Item.DisplayName, $Item.Name, $Item.ServerName, $Item.FeatureType)
			$Save = "$($Item.DisplayName)$($Item.Name)"
			if ($First)
			{
				$First = $False
			}
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix N Windows Installed Components' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function OutputAppendixO
{
	#added in V1.16
	Write-Verbose -Message ('{0}: Create Appendix O PVS Processes' -f (Get-Date -UFormat "%F %r (%Z)"))

	$Script:PVSProcessItems = $Script:PVSProcessItems | Sort-Object ProcessName, ServerName
	
	if ($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_Assessment_AppendixO_PVSProcesses.csv"
		$Script:PVSProcessItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Write-Output -Message "Appendix O - PVS Processes"
	Line 0 ''
	Write-Output -Message "Process Name  Server Name     Status     "
	Write-Output -Message "========================================="
	#       1234567890123S123456789012345S12345678901
	#       StreamProcess XXXXXXXXXXXXXXX Not Running
	#       13            15              11
	$Save = ''
	$First = $True
	if ($Script:PVSProcessItems)
	{
		ForEach($Item in $Script:PVSProcessItems)
		{
			if (!$First -and $Save -ne "$($Item.ProcessName)")
			{
				Line 0 ''
			}

			Line 1 ( "{0,-13} {1,-15} {2,-11}" -f `
			$Item.ProcessName, $Item.ServerName, $Item.Status)
			$Save = "$($Item.ProcessName)"
			if ($First)
			{
				$First = $False
			}
		}
	} else {
		Write-Output -Message "<None found>"
	}
	Line 0 ''

	Write-Verbose -Message ('{0}: Finished Creating Appendix O PVS Processes' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

Function ShowScriptOptions
{
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: AdminAddress       : $($AdminAddress)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: CSV                : $($CSV)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Domain             : $($Domain)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Filename1          : $($Script:filename1)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Folder             : $($Script:pwdpath)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: From               : $($From)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: PVS Farm Name      : $($Script:farm.farmName)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: PVS Version        : $($Script:PVSFullVersion)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Smtp Port          : $($SmtpPort)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Smtp Server        : $($SmtpServer)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Title              : $($Script:Title)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: To                 : $($To)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Use SSL            : $($UseSSL)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: User               : $($User)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: OS Detected        : $($Script:RunningOS)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: PoSH version       : $($Host.Version)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: PSCulture          : $($PSCulture)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: PSUICulture        : $($PSUICulture)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: Script start       : $($Script:StartTime)' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
	Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))
}

#script begins

$script:startTime = get-date

# v1.17 - switch to using a StringBuilder for $global:Output
[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

Write-Verbose -Message ('{0}: Checking for McliPSSnapin' -f (Get-Date -UFormat "%F %r (%Z)"))
if (!(Check-NeededPSSnapins "McliPSSnapIn")){
	#We're missing Citrix Snapins that we need
	Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Script will now close."
	Exit
}

SetupRemoting

VerifyPVSServices

GetPVSVersion

GetPVSFarm

ShowScriptOptions

ProcessPVSFarm

ProcessPVSSite

ProcessvDisksinFarm

ProcessStores

OutputAppendixA

OutputAppendixB

OutputAppendixC

OutputAppendixD

OutputAppendixE

OutputAppendixF

OutputAppendixF2

OutputAppendixG

OutputAppendixH

#outputs AppendixI
ProcessvDisksWithNoAssociation

OutputAppendixJ

OutputAppendixK

OutputAppendixL

OutputAppendixM

OutputAppendixN

OutputAppendixO

Write-Verbose -Message ('{0}: Finishing up document' -f (Get-Date -UFormat "%F %r (%Z)"))
#end of document processing

SaveandCloseTextDocument

Write-Verbose -Message ('{0}: Script has completed' -f (Get-Date -UFormat "%F %r (%Z)"))
Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))

$GotFile = $False

if (Test-Path "$($Script:FileName1)")
{
	Write-Verbose -Message ('{0}: $($Script:FileName1) is ready for use' -f (Get-Date -UFormat "%F %r (%Z)"))
	$GotFile = $True
}
Else
{
	Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
	Write-Error "Unable to save the output file, $($Script:FileName1)"
}

#email output file if requested
if ($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
{
	$emailAttachment = $Script:FileName1

	SendEmail $emailAttachment
}

Write-Verbose -Message ('{0}: ' -f (Get-Date -UFormat "%F %r (%Z)"))

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose -Message ('{0}: Script started: $($Script:StartTime)' -f (Get-Date -UFormat "%F %r (%Z)"))
Write-Verbose -Message ('{0}: Script ended: $(Get-Date)' -f (Get-Date -UFormat "%F %r (%Z)"))
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose -Message ('{0}: Elapsed time: $($Str)' -f (Get-Date -UFormat "%F %r (%Z)"))
$runtime = $Null
$Str = $Null

<#
	.SYNOPSIS
		Creates a basic assessment of a Citrix PVS 5.x or later farm.
	.DESCRIPTION
		Creates a basic assessment of a Citrix PVS 5.x or later farm.
		
		Creates a text document named after the PVS farm.
		
		Register the old string-based PVS Console PowerShell Snap-in.

		For versions of Windows prior to Windows 8 and Server 2012, run:
		
		For 32-bit:
			%systemroot%\Microsoft.NET\Framework\v2.0.50727\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
		For 64-bit:
			%systemroot%\Microsoft.NET\Framework64\v2.0.50727\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"

		For Windows 8.1 and later, Server 2012 and later, run:
		
		For 32-bit:
			%systemroot%\Microsoft.NET\Framework\v4.0.30319\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
		For 64-bit:
			%systemroot%\Microsoft.NET\Framework64\v4.0.30319\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"

		All lines are one line. 

		If you are running 64-bit Windows, you will need to run both commands so 
		the snap-in is registered for both 32-bit and 64-bit PowerShell.
		
		NOTE: The account used to run this script must have at least Read access to the SQL 
		Server that holds the Citrix Provisioning databases.

	.PARAMETER AdminAddress
		Specifies the name of a PVS server that the PowerShell script will connect to. 
		Using this parameter requires the script be run from an elevated PowerShell session.
		Starting with V1.11 of the script, this requirement is now checked.
		This parameter has an alias of AA
	.PARAMETER CSV
		Will create a CSV file for each Appendix.
		The default value is False.
		
		Output CSV filename is in the format:
		
		PVSFarmName_Assessment_Appendix#_NameOfAppendix.csv
		
		For example:
			TNPVSFarm_Assessment_AppendixA_AdvancedServerItems1.csv
			TNPVSFarm_Assessment_AppendixB_AdvancedServerItems2.csv
			TNPVSFarm_Assessment_AppendixC_ConfigWizardItems.csv
			TNPVSFarm_Assessment_AppendixD_ServerBootstrapItems.csv
			TNPVSFarm_Assessment_AppendixE_DisableTaskOffloadSetting.csv	
			TNPVSFarm_Assessment_AppendixF_PVSServices.csv
			TNPVSFarm_Assessment_AppendixG_vDiskstoMerge.csv	
			TNPVSFarm_Assessment_AppendixH_EmptyDeviceCollections.csv	
			TNPVSFarm_Assessment_AppendixI_UnassociatedvDisks.csv	
			TNPVSFarm_Assessment_AppendixJ_BadStreamingIPAddresses.csv	
			TNPVSFarm_Assessment_AppendixK_MiscRegistryItems.csv
			TNPVSFarm_Assessment_AppendixL_vDisksConfiguredforServerSideCaching.csv	
			TNPVSFarm_Assessment_AppendixM_MicrosoftHotfixesandUpdates.csv
			TNPVSFarm_Assessment_AppendixN_InstalledRolesandFeatures.csv
			TNPVSFarm_Assessment_AppendixO_PVSProcesses.csv
	.PARAMETER User
		Specifies the user used for the AdminAddress connection. 
	.PARAMETER Domain
		Specifies the domain used for the AdminAddress connection. 
	.PARAMETER Password
		Specifies the password used for the AdminAddress connection. 
	.PARAMETER Folder
		Specifies the optional output folder to save the output report. 
	.PARAMETER SmtpServer
		Specifies the optional email server to send the output report. 
	.PARAMETER SmtpPort
		Specifies the SMTP port. 
		The default port is 25.
	.PARAMETER UseSSL
		Specifies whether to use SSL for the SmtpServer.
		THe default is False.
	.PARAMETER From
		Specifies the username for the From email address.
		If SmtpServer is used, this is a required parameter.
	.PARAMETER To
		Specifies the username for the To email address.
		If SmtpServer is used, this is a required parameter.
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1
		
		Will use all Default values.
		LocalHost for AdminAddress.
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -AdminAddress PVS1 -User cwebster -Domain WebstersLab -Password Abc123!@#

		This example is usually used to run the script against a PVS Farm in 
		another domain or forest.
		
		Will use:
			PVS1 for AdminAddress.
			cwebster for User.
			WebstersLab for Domain.
			Abc123!@# for Password.
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -AdminAddress PVS1 -User cwebster

		Will use:
			PVS1 for AdminAddress.
			cwebster for User.
			Script will prompt for the Domain and Password
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -Folder \\FileServer\ShareName
		
		Output file will be saved in the path \\FileServer\ShareName
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -SmtpServer mail.domain.tld -From 
		XDAdmin@domain.tld -To ITGroup@domain.tld -ComputerName DHCPServer01
		
		Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
		sending to ITGroup@domain.tld.
		If the current user's credentials are not valid to send email, the user will be prompted 
		to enter valid credentials.
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
		-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com
		
		Script will use the email server smtp.office365.com on port 587 using SSL, sending from 
		webster@carlwebster.com, sending to ITGroup@carlwebster.com.
		If the current user's credentials are not valid to send email, the user will be prompted 
		to enter valid credentials.
	.EXAMPLE
		PS C:\PSScript > .\PVS_Assessment.ps1 -CSV
		
		Will use all Default values.
		LocalHost for AdminAddress.
		Creates a CSV file for each Appendix.
	.INPUTS
		None.  You cannot pipe objects to this script.
	.OUTPUTS
		No objects are output from this script.  This script creates a text file.
	.NOTES
		NAME: PVS_Assessment.ps1
		VERSION: 1.20
		AUTHOR: Carl Webster, Sr. Solutions Architect at Choice Solutions (with a lot of help from BG a, now former, Citrix dev)
		LASTEDIT: July 8, 2019
#>
