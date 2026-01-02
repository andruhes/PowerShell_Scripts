## Nanoweb ##
# import nano server ps bits into session 
impoprt-module e:\NanoServer\NanoServerImageGenerator -verbose

# Create a nanoserver web server
new-nanoserverimage -mediapath e:\Nano_Web -targetpath D:\Nano_Web\nanoweb.vhdx -deploymenttype guest -edition standard -computername nanoweb -interfacenameorindex ethernet -ipv4address 192.168.101.32 -ipv4subnetmask 255.255.255.0 -ipv4gateway 192.168.100.1 -ipv4dns ("192.168.100.43" , "8.8.8.8") -package Microsoft-NanoServer-IIS-Package -administratopassword (ConvertTo-SecureString -String "Liixo$01" -AsPlainText -Force) 

#View available packages
Get-nanoserverpackage -mediapath e:

# Create and start vm
New-VM -name nanoweb vhdpath D:\Nano_Web\nanoweb.vhdx -memorystartupbytes 1gb -generation 2 | start-vm

# Ps Session to Nanoweb
Enter-PSSession -VMName nanoweb

#View IIS service, logs and files
Get-Service w3svc
Get-WinEvent -ListLog microsoft-iis*
Get-ChildItem C:\inetpub

# View Installed packages
Get-WindowsPackage -online

# Disable firewall
Set-NetFirewallProfile -name public, private, domain -Enabled False

#Exit session 
Exit-PSSession
