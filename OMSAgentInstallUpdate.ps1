##########################################################################################
#                                                                                        #
# Install, Update and reconfigure OMS Agent to SCOM 2016 Version, and add new workspaces #
                                                                  #
#                                                                                        #
# Filename: OMSInstallUpdate.ps1                                                         #
#                                                                                        #
# Version: 1.0                                                                           #
#                                                                                        #
# Notes: All examples are provided “AS IS” with no warranty expressed or implied.        # 
#        Run at your own risk.                                                           #
#                                                                                        #
##########################################################################################


#Configuration Instructions and Explanation for Each Variable.

#All variables that could be configured are marked with **** You should modify this one as needed**** at the end of the line.

#This script Installs or updates the MMA agent from the public URL, configure the OMS Workspaces as needed, and can modify the SCOM Managment Group and OMS Proxy Settings.


###############################################################################################################################################################


<#

#Configuration Flags

$AddManagementGroup - If set to 0 the MG configuration block from the script will be ignored, if set to 1 it will apply the MG configuration to the agent, currently support adding just one MG. **** You should modify this one as needed****

$ClearManagementGroup - If set to 0 all Management Groups in the agent will be cleared, if set to 1 it will add the new Management Group to the agent. **** You should modify this one as needed****

$AddOMSProxy - If set to 0 the OMS proxy configuration block from the script will be ignored, if set to 1 it will apply the OMS proxy configuration to the agent. **** You should modify this one as needed****

$OMSAddWS = If set to 0 the OMS add workspaces configuration block from the script will be ignored, if set to 1 it will add the specified OMS Workspaces. **** You should modify this one as needed****

$OMSRemoveOldWS = If set to 0 it wonpt remove any current OMS workspaces in the agent, if set to 1 it remove all current OMS workspaces in the agent. **** You should modify this one as needed****


#Management Group Information

*****This configuration is only applied If you specified $AddManagementGroup to 1*****

$MGName - This is the Management Group name that will be applied to the agent. **** You should modify this one as needed****

$MSName - This is the Management Server name that is hosting the Management Group and will be applied to the agent. **** You should modify this one as needed****

$MSPort - This is the port for communicating with the Management Server that will be applied to the agent, default is 5723. ****This configuration is optional****


#OMS Proxy Information

*****This configuration is only applied If you specified $AddOMSProxy to 1*****

$ProxyURL - The URL of the proxy you want to use, for example: "http://proxy.test.com:80". **** You should modify this one as needed****

$proxyuser - The user that will be used to authenticate to the proxy server, for example: "YourUserHere". **** You should modify this one as needed****

$proxypass - The passworrd for $proxyuser stored in a secure way, for example : ConvertTo-SecureString -String "YourPasswordHere" -AsPlainText -Force.

$proxycred - It assigns $proxyuser and $proxypass to a credential variable.

$proxyuser - Assigns the $proxycred user to use it when loading configuration for proxy.


#Download variables: This are used to specify the download url and the file local file path.

$OMSLocalPath - The directory you want to use to download the agent .exe file, modify to the path of your preference. **** You should modify this one as needed****

$OMSInstallerPath - The full path to the file to download, incuding the filename, you don´t need to modify this one. **** You should modify this one as needed****

$OMSDownloadURL - The public URL to download the agent, you just need to modify this if there´s a new version of the agent release to a new URL.


#Install command for the MMA executable.

$OMSInstallerParameters - This are the command line installation intructions.


#Get current time and date to determine the running time of the script.

$start_time - Gets the date to report how long the script took to run.


#Assign OMS Workspaces: Assign IDs and Keys in two different arrays to definr the agent configuration.

$OMSWorkspaceId - The IDs of the workspaces you want to add, the should be in ordes separated by a comma, for example: OMSWS1, OMSWS2, OMSWS3... **** You should modify this one as needed****

$OMSWorkspaceKey - The Keys of the workspaces you want to add, this should be in the same order of the workspaces for example: OMSKey1, OMSKey2, OMSKey3... **** You should modify this one as needed****


#Mail Parameters

$body - Base initial body for the email recommendation. **** You should modify this one as needed****

$user - The username used to send the email, for example: "monitoring@outlook.com". **** You should modify this one as needed****

$pass - Password for the user used to send the email. **** You should modify this one as needed****

$cred - Adds the user and password to a credential variable that will be use to authenticate with the mail service.

$sender - The sender's email, for example "admin@outlook.com". **** You should modify this one as needed****

$recipient - The recipient for the message, for example: "User <user@outlook.com>". **** You should modify this one as needed****

$smtp - Smtp address for the mail service for example "smtp.outlook.com". **** You should modify this one as needed****

$port - port used to connect to the smtp, usually 587. **** You should modify this one as needed****


#Check for previous installation and version

$OMSAgent = Gets the registry key for the Microsoft Monitoring Agent Installation

$OMSInstalled = Gets 1 if the agent is installed and 0 if not installed.

#>


###############################################################################################################################################################


#Configuration Flags

$AddManagementGroup = 0

$ClearManagementGroup = 0

$AddOMSProxy = 0

$OMSAddWS = 1

$OMSRemoveOldWS = 1


#Management Group Information

$MGName = "ManagementGroupName"

$MSName = "ManagementServerName"

$MSPort = 5723


#OMS Proxy Information

$ProxyURL = "http://your.proxy.com:80" 

$Proxyuser = "ProxyUser"

$Proxypass = ConvertTo-SecureString -String "YourPassword" -AsPlainText -Force

$Proxycred = New-Object System.Management.Automation.PSCredential $Proxyuser, $Proxypass

$Proxyuser = $Proxycred.UserName


#Download variables

$OMSLocalpath = "C:\Temp"

$OMSInstallerPath = "$OMSLocalpath\MMASetup-AMD64.exe"

$OMSDownloadURL = "http://download.microsoft.com/download/0/C/0/0C072D6E-F418-4AD4-BCB2-A362624F400A/MMASetup-AMD64.exe"


#Install command for the MMA executable.

$OMSInstallerParameters = '/Q:A /R:N /C:"setup.exe /qn AcceptEndUserLicenseAgreement=1"'


#Get current time and date to determine the running time of the script.

$start_time = Get-Date


#Assign OMS Workspaces: Assign IDs and Keys in two different arrays to definr the agent configuration.

$OMSWorkspaceId = "OMSWS1, OMSWS2, OMSWS3"

$OMSWorkspaceKey = "OMSKey1, OMSKey2, OMSKey3"


#Mail Parameters

$body = "The server $env:computername has been updated with OMS Agent version 134228607 and the next specified configurations:"

$user = "user@outlook.com"

$pass = ConvertTo-SecureString -String "yourpassword" -AsPlainText -Force

$cred = New-Object System.Management.Automation.PSCredential $user, $pass

$sender = "Me <user@outlook.com>"

$recipient = "user@outlook.com"

$smtp = "smtp.outlook.com"

$port = 587


#Check for previous installation and version

$OMSAgent = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq 'Microsoft Monitoring Agent'}

$OMSInstalled = @($OMSAgent).Count


if ($OMSInstalled -eq 1){

    
    #Agent version validation

    $OMSVersion = $OMSAgent.Version

    $OMSUpToDate = $OMSVersion -eq 134228607

    
    if ($OMSUpToDate -eq 0){
    
        Write-Host "Update Agent"

        Write-Host "Downloading Agent"

        #Download Agent

        Import-Module BitsTransfer

        Start-BitsTransfer -Source $OMSDownloadURL -Destination $OMSInstallerPath

        Write-Host "Updating"

        #Update Agent

        $setup = Start-Process "$OMSInstallerPath" -ArgumentList "$OMSInstallerParameters" -PassThru -Wait


        if($setup.ExitCode -eq 0){

            write-Host "Successfully Updated"

        }

        
        else{

            write-Host "Agent installation error with code" $setup.StandardError

            Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

            Exit

        }


        #Apply configuration to OMS Workspaces in the agent.
        
        if($OMSAddWS -eq 1 -or $OMSRemoveOldWS -eq 1) {
        
            $OMSWS = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'

        
            #Remove All Existing OMS Workspaces
        
        
            if($OMSRemoveOldWS -eq 1){

                $OMSRemoveWS = $OMSWS.GetCloudWorkspaces()


                foreach ($WS in $OMSRemoveWS) {
        
                    Write-Host "Removing Old Workspaces"            
                
                    $OMSWS.RemoveCloudWorkspace($WS.workspaceId)

                    $body += "`n `nThe Workspace", $WS.workspaceId, "was Successfully Removed."

                }

            }

        
            #Assign New OMS Workspaces


            if($OMSAddWS -eq 1){

            
                for ($i = 0;$i -lt $OMSWorkspaceId.Count; $i ++) {

                    Write-Host "Assigning Workspaces"

                    $OMSWS.AddCloudWorkspace($OMSWorkspaceId[$i], $OMSWorkspaceKey[$i])
                    
                    $body += "`n `nThe Workspace", $OMSWorkspaceId[$i], "was Successfully Added."    
        
                }


            }    


}


        #Add management Group if Required

        
        if($AddManagementGroup -eq 1 -or $ClearManagementGroup -eq 1){
            
            
            if($ClearManagementGroup -eq 1){
            
                $OMSMG = $OMSWS.GetManagementGroups()

                foreach($MG in $OMSMG){
                    
                    Write-Output "Removing old Management Group" $MG.managementGroupName
                    
                    $OMSWS.RemoveManagementGroup($MG.managementGroupName)

                }


            }
            
            
            if($AddManagementGroup -eq 1){

                Write-Output "Setting Management Group to ${MGName} in hosted Management Server ${MSName} and port ${MSPort}"
            
                $OMSWS.AddManagementGroup($MGName, $MSName, $MSPort)

                $body += "`n `nThe server was added to the Management Group $MGName hosted in Server $MSName with port $MsPort."

            }


        }
        

        #Add OMS Proxy if Required

        
        if($AddOMSProxy -eq 1){

            Write-Output "Clearing proxy settings."
            
            $OMSWS.SetProxyInfo('', '', '')
            
            Write-Output "Setting Proxy to $ProxyURL with proxy username of $ProxyUser."

            $OMSWS.SetProxyInfo($ProxyURL, $ProxyUser, $Proxycred.GetNetworkCredential().password)

            $body += "`n `nThe OMS proxy $ProxyURL was added with user $Proxyuser."

        }


        #Reload agent configuration to update settings with workspaces added and removed       

        $OMSWS.ReloadConfiguration()
        
        $OMSWS.GetManagementGroups()

        write-Host "Workspaces Updated"

        }


    else {

        #Apply configuration to OMS Workspaces in the agent.
        
        if($OMSAddWS -eq 1 -or $OMSRemoveOldWS -eq 1) {
        
            $OMSWS = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'

        
            #Remove All Existing OMS Workspaces
        
        
            if($OMSRemoveOldWS -eq 1){

                $OMSRemoveWS = $OMSWS.GetCloudWorkspaces()


                foreach ($WS in $OMSRemoveWS) {
        
                    Write-Host "Removing Old Workspaces"            
                
                    $OMSWS.RemoveCloudWorkspace($WS.workspaceId)

                    $body += "`n `nThe Workspace", $WS.workspaceId, "was Successfully Removed."

                }

            }

        
            #Assign New OMS Workspaces


            if($OMSAddWS -eq 1){

            
                for ($i = 0;$i -lt $OMSWorkspaceId.Count; $i ++) {

                    Write-Host "Assigning Workspaces"

                    $OMSWS.AddCloudWorkspace($OMSWorkspaceId[$i], $OMSWorkspaceKey[$i])
                    
                    $body += "`n `nThe Workspace", $OMSWorkspaceId[$i], "was Successfully Added."    
        
                }


            }    


}


        #Add management Group if Required

        
        if($AddManagementGroup -eq 1 -or $ClearManagementGroup -eq 1){
            
            
            if($ClearManagementGroup -eq 1){
            
                $OMSMG = $OMSWS.GetManagementGroups()

                foreach($MG in $OMSMG){
                    
                    Write-Output "Removing old Management Group" $MG.managementGroupName
                    
                    $OMSWS.RemoveManagementGroup($MG.managementGroupName)

                }


            }
            
            
            if($AddManagementGroup -eq 1){

                Write-Output "Setting Management Group to ${MGName} in hosted Management Server ${MSName} and port ${MSPort}"
            
                $OMSWS.AddManagementGroup($MGName, $MSName, $MSPort)

                $body += "`n `nThe server was added to the Management Group $MGName hosted in Server $MSName with port $MsPort."

            }


        }
        

        #Add OMS Proxy if Required

        
        if($AddOMSProxy -eq 1){

            Write-Output "Clearing proxy settings."
            
            $OMSWS.SetProxyInfo('', '', '')
            
            Write-Output "Setting Proxy to $ProxyURL with proxy username of $ProxyUser)."

            $OMSWS.SetProxyInfo($ProxyURL, $ProxyUser, $Proxycred.GetNetworkCredential().password)

            $body += "`n `nThe OMS proxy $ProxyURL was added with user $Proxyuser."

        }

        
        #Reload agent configuration to update settings with workspaces added and removed       

        $OMSWS.ReloadConfiguration()

        $OMSWS.GetManagementGroups()

        write-Host "Workspaces Updated"

    }


}


else{

    Write-Host "Install Agent"

    
    #Download Agent

    Write-Host "Downloading Agent"

    Import-Module BitsTransfer

    Start-BitsTransfer -Source $OMSDownloadURL -Destination $OMSInstallerPath

    Write-Host "Installing Agent"

    
    #Install Agent

    $setup=Start-Process "$OMSInstallerPath" -ArgumentList "$OMSInstallerParameters" -PassThru -Wait

    
    if($setup.ExitCode -eq 0){

        write-Host "Successfully Installed"

    }


    else{

        write-Host "Agent installation error with code" $setup.StandardError

        Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

    Exit

    }
 

    #Apply configuration to OMS Workspaces in the agent.
        
    
    if($OMSAddWS -eq 1) {
        
        $OMSWS = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'

 
        #Assign New OMS Workspaces

            
        for ($i = 0;$i -lt $OMSWorkspaceId.Count; $i ++) {

            Write-Host "Assigning Workspaces"

            $OMSWS.AddCloudWorkspace($OMSWorkspaceId[$i], $OMSWorkspaceKey[$i])
            
            $body += "`n `nThe Workspace", $OMSWorkspaceId[$i], "was Successfully Added."    
        
        }  


    }


    #Add management Group if Required


    if($AddManagementGroup -eq 1 -or $ClearManagementGroup -eq 1){
            
            
        if($ClearManagementGroup -eq 1){
            
            $OMSMG = $OMSWS.GetManagementGroups()

            foreach($MG in $OMSMG){
                    
                Write-Output "Removing old Management Group" $MG.managementGroupName
                    
                $OMSWS.RemoveManagementGroup($MG.managementGroupName)

            }


        }
            
            
        if($AddManagementGroup -eq 1){

            Write-Output "Setting Management Group to ${MGName} in hosted Management Server ${MSName} and port ${MSPort}"
            
            $OMSWS.AddManagementGroup($MGName, $MSName, $MSPort)

            $body += "`n `nThe server was added to the Management Group $MGName hosted in Server $MSName with port $MsPort."

        }


    }
        

    #Add OMS Proxy if Required


    if($AddOMSProxy -eq 1){

        Write-Output "Clearing proxy settings."
            
        $OMSWS.SetProxyInfo('', '', '')
            
        Write-Output "Setting Proxy to $ProxyURL with proxy username of $Proxyuser."

        $OMSWS.SetProxyInfo($ProxyURL, $Proxyuser, $Proxypass)

        $body += "`n `nThe OMS proxy $ProxyURL was added with user $Proxyuser."

    }

        
    #Reload agent configuration to update settings with workspaces added and removed       

    $OMSWS.ReloadConfiguration()

    $OMSWS.GetManagementGroups()

    write-Host "Workspaces Updated"

    write-Host "Agent Installation and Config was Successful"

}


#Send Email

$body += "`n `nTime taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

write-Host "Sending Email"

Send-MailMessage -from $sender -to $recipient -subject "OMS Installed/Updated in $env:computername" -body $body -dno onSuccess, onFailure -smtpServer $smtp -Port $port -Credential $cred -UseSsl


#Finished overall time output

Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"