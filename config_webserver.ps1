#create a DSC configuration to install IIS and support remote management.
Configuration IISConfig {
        
    #define input parameter
    param( 
    [string[]] $computername = 'localhost'
    )
     #target machine(s) based on input param
     node $computername {
       
       #configure the LCM
       LocalConfigurationManager {
       ConfigurationMode = "ApplyAndAutoCorrect"
       ConfigurationModeFrequencyMins = 15
       RefreshMode = "Push"
       }
       
       #install the IIS server role
       WindowsFeature IIS {
       Ensure = "Present"
       Name = "web-server"
       }

       #install the IIS remote management service 
       WindowsFeature IISManagement {
       Name = 'Web-Mgmt-Service'
       Ensure = 'Present'
       DependsOn = @('[WindowsFeature]IIS')
       }

       #enable IIS remote management
       Registry RemoteManagement {
       Key = 'HKLM:\SOFTWARE\Microsoft\WebManagement\Server'
       ValueName = 'EnableRemoteManagement'
       ValueType = 'Dword'
       ValueData = '1'
       DependsOn = @('[WindowsFeature]IIS','[WindowsFeature]IISManagement')
       }

       #configure remote management service 
       Service WMSVC {
       Name = 'WMSVC'
       StartupType = 'Automatic'
       State = 'Running'
       DependsOn = '[Registry]RemoteManagement'
       }
    }
}

IISConfig -computername server2016core1 -OutputPath C:\IT.info

Start-DscConfiguration -Path C:\IT.info -Wait -Verbose
