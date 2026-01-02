#Install Net framework
Get-windowsFeature net-framework-core
Install-windowsfeature net-framework-core -source D:\source\sxs

#register the smt tools
C:\Migrationtools\SMT_WS12R2_amd64\Smigdeploy.exe

#import smig cmdlets into session
Add-PSSnapin microsoft.windows.servermanager.migration
Get-Command -Module "migration"

# Stop DHCP
Stop-service -Name DHCPServer

# View roles and features eligible for migration
get=smigserverfeature

#Create the migration store
export-smigserversetting -featureID DHCP -path C:\Migrationtools -verbose

#Copy migration store to destination server
Copy-Item -path C:\Migrationtools -Destination \\srvgui\Migrationtools -Recurse
