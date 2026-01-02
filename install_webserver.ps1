#Enter a powershell session on a remore machine
Enter-PSSession -ComputerName server2016core1

#view all available features
Get-WindowsFeature -name web*

#install web server (iis) role and management service 
Install-WindowsFeature -Name web-server, Web-Mgmt-Service

#view installed feachers
Get-WindowsFeature | Where-Object Installed -eq True

# configure remote management for IIS
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WebManagement\Server" -Name "EnableRemoteManagement" -Value 1

#configure remote management service to start automatically
Set-Service WMSVC -StartupType Automatic

#exit remote powershell session
Exit-PSSession

# sent a command over to a remote machine 
Invoke-Command -ComputerName server2016core1 -ScriptBlock {Get-Service W3SVC, WMSVC}

#get list of services
Get-Service -ComputerName server2016core1
