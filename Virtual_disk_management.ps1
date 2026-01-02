#Get windows feature
get-windowsfeature *hyper-v*

#Install Hyper-v powershell cmdlets
Install-WindowsFeature -name hyper-v -IncludeManagementTools

#View VHD related cmdlets 
Get-Command -Module hyper-v name *VHD*

#Create a new dynamic vhdx
new-vhd -path V:\Virtuail Disks\virtual disk 2.vhdx -sizebytes 5gb -dynamic

#Mount it
mount-vhd -path V:\Virtual disks\virtual disk 2.vhdx

#Initialize? partition and format it
initialize-disk -number 5 -PartitionStyle GPT -PassThru | New-Partition -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS

#unmount it
dismount-vhd -path V:\Virtual Disks\virtual disk 2.vhdx

#View disk information
get-vhd -path V:\Virtual Disks\virtual disk 2.vhdx