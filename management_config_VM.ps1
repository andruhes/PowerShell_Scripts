### Dynamic memory###

#Configure dynamic memory
Get-VM | Set-VM -MemoryStartBytes 2gb -DynamicMemory -MemoryMinimumBytes 1gb -MmeoryMaximumBytes 3gb

#Configure buffer and weight
Set-VMMemory -VmName serverhyperv2 -Buffer 10 -Priority 100

## NUMA ##

#view Numa topology
Get-VMHostNumaNode

# Disable NUMA
Set-VMHost -NumaSpanningEnable $false

## Smart paging ##
Get-VM | Set-VM -SmartpagingFilePath D:\VMs\Paging

## Resource metering ##

#Enable resource metering 
Get-VM | Enable-VMResourceMetering

#View usage stats
Get-VM | Measure-VM

#Configure collection intervals
Set-VMHost -ResourceMeteringSaveInterval 00:01:00

# Create a resource pool to measure the memory
New-VMResourcePool -name VMPOOL -ResourcePoolType Memory

# Add VMs memory to Pool
Set-VMMemory -ResourcePoolName VMPOOL -VMName serverhyperv2

#View pool usage stats
Measure-VMResourcePool -Name VMPOOL | fl "mem"

# Disable resource metering 
Get-VM | Diasble-VMResourceMetering

# Hot add/remove memory
Get-VM | Set-VMMemory -StartupBytes 3072mb

#Enable Enhanced session mode on the host
Set-VMHost -EnableEnhancedSession $true

#Enable secure boot
Get-VM | Set-VMFirmware -EnableSecureBoot On _secureBootTemplate MicrosoftWindows
Get-VM | Get-VMFirmware


# View integration services
Get-VMInttegrationService -VMName serverhyperv2

# Enable guest services
Get-VMInttegrationService -VMName serverhyperv2 -name "Guest service interface"

#Export a VM
Export-VM -Name Test -Path D:\Exports

# Import a VM
Import-VM -path 'D:\Exports\Test\'WAY'.vmcx'

# Upgarde VM version
Get-VM | Update-VMVersion

