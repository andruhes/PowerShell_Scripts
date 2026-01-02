#import nano server ps bits into session
Import-Module E:\NanoServer\NanoServerImageGenerator -verbose

##nanosserver##
#create nano server vhdx
new-nanoserverimage -mediaoath e: -basepath d:\Nano_Server -targetpath d:\Nano_Server\nanoserver.vhdx -deploymenttype guest -edition standard -computername nanoserver -administratorpassword (ConvertTo-SecureString -String 'Lixio$01' -AsPlainText -Force)

#create and start vm
New-VM -name nanoserver -vhdpath d:\Nano_Server\nanoserver.vhdx -memorystartupbytes 1gb -generation 2 | start-vm

#PS Session to NanoServer
Enter-PSSession -VMName nanoserver

#Look at directory Get-Childtem c:
cd \

# View prosesses, services, event logs
Get-Process
Get-Service
Get-WinEvent

#Exit session
Exit-PSSession
