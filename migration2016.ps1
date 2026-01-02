#view server magation tools feature 
Get-Windowsfeature -name megration

#install windows server migration tools
Install-windowsfeature -Name migration

#navigate to set directory
cd \windows\system32\servermigrationtools

#generate package for source machine
.\smigdeploy.exe /package /architecture amd64 /os WS12R2 /path C:\Migrationtools

#copy package over to source machine
Copy-Item =path C:\Migrationtools\SMT_ws12R2_amd64 -destination \\server2012gui\Migrationtools -Recurse

#Переходим к серверу откуда будет проводится миграция 

#import smig cmdlets into session
Add-PSSnapin microsoft.windows.servermanager.migration

#install DHCP role and import settings
Import-smigserversetting -featureID DHCP -path C:\Migrationtools\Migrationtools -verbose -force

