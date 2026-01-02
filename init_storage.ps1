#view storege related cmdlets
get-command -module storage

#view disks
Get-Disk

#initialize disk as GPT 
initialize-disk -Number 2 -PartitionStyle GPT

#view partitions
Get-Partition

# partition an entire disk
New-Partition -DiskNumber 2 -UseMaximumSize -DriveLetter I

#view voluems 
Get-Volume

# Format with a file system
Format-Volume -DriveLetter I -FileSystem NTFS -AllocationUnitSize 4096 -NewFileSystemLabel "IT DAta"

# formate remaining disk with a single command
Get-Disk
    Where-Object partitionstyle -EQ "RAW"
     intialize-disk -partitionstyle GPT -passthru
      New-Partition -UseMaximumSize -DriveLetter V
       Format-Volume -FileSystem NTFS -AllocationUnitSize 65536 -NewFileSystemLabel "VM DATA"
