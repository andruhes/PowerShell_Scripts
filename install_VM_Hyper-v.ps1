#Create generation 2 VM 
new-VM -name serverhyperv2 -MemorystartupBytes 1gb -NewVHDPath D:\VMs\serverhyperv2.vhdx -NewVHDSizeBytes 30gb -generation 2

#Add SCSI controller for DVD drive
Add-VMScsiController -VMName serverhyperv2

#Add DVD drive to SCSI controller with Windows Server 2016 installation mounted
Add-VMDvdDrive -VMName serverhyperv2 _controller Number 1 -ControllerLocation 0 -path D:\Images\en_windows_server_2016updated_feb_2018_x64_dvd_11636692.iso

#Start the VM 
Start-VM -nmae serverhyperv2