<#
.SYNOPSIS
   <Connect your session to remote O365 Powershell and run common powershell tasks>
.DESCRIPTION
   <Powershell 5.0  (Windows 10)>
.PARAMETER <paramName>
   <Collection of usefull queries towards O365 tenant>
.EXAMPLE
   <Use one of the premade commands to list info as wanted.
   msolinfo | ft >
.AUTHOR
        Not a programmer, so I know the coding isnt optimal or as clean as it probably could be ;(
        tore@cegal.com ;)
#>
<#Global variables for use in functions and outside functions#>
        $Global:LiveCred
        $Global:Company
        $Global:Manual
        $Global:A
        $Global:B
       
<# Functions #>
        Function Menu {
                        <#Used this resource to create the menu: https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/ #>
                        Get-PSSession | ft
                        [int]$xMenuChoiceA = 0
                            while ( $xMenuChoiceA -lt 1 -or $xMenuChoiceA -gt 6 ){
                            Write-host "---------------------------------------------------------------------------------------"
                            Write-host "Select a partner tenant to connect your Powershell session to (Or enter your own)"
                            Write-host "---------------------------------------------------------------------------------------"
                            Write-host "1. Partner #1"
                            Write-host "2. Partner #2"
                            Write-host "3. Partner #3"
                            Write-host "4. Partner #4"
                            Write-host "5. Enter your own Tenant"
                            Write-host "---------------------------------------------------------------------------------------"
                            Write-host "6. Quit and exit"
                            Write-host "---------------------------------------------------------------------------------------"
                        [Int]$xMenuChoiceA = read-host "Please enter an option 1 to 6..." }
                        Switch( $xMenuChoiceA ){
                            1{$global:company = "admin@partner1.onmicrosoft.com"}
                            2{$global:company = "admin@partner2.onmicrosoft.com"}
                            3{$global:company = "admin@partner3.onmicrosoft.com"}
                            4{$global:company = "admin@partner4.onmicrosoft.com"}
                            5{ $Global:Manual = Read-Host -Prompt "Type in your Tenant username:"
                               $global:company = $Global:manual
                               }
                            6{Exit}
                        default{}
                    }
                    O365
}
        Function TenantMenu {
                Get-PSSession | ft
                Write-Host " "
                Write-Host "---------------------------------------------------------------------"
                Write-Host "Aliases available:('Tenantmenu' to see this again, 'Menu' to connect to new O365 session)" -ForegroundColor Yellow
                Write-Host "Logged in user:" $global:company
                Write-Host "Use following commands to administer your services: Msonline, Skype4B, Sharepoint, Compliance" -ForegroundColor Yellow
                Write-Host "---------------------------------------------------------------------"
                Write-Host "1. 'Mailbox' (Lists Displayname, Windows Emailaddress, Primary SMTP)" -ForegroundColor Yellow
                Write-Host "2. 'Licensed' (Lists users with valid license, run msonline first)" -ForegroundColor Yellow
                Write-Host "3. 'NotLicensed' (Lists users without license, run msonline first)" -ForegroundColor Yellow
                Write-Host "4. 'Mailboxsize' (Lists users and mailboxsize in KB & GB)" -ForegroundColor Yellow
                Write-Host "5. 'Aliases' (Lists Displayname, WindowsEmailaddress, PrimarySmtpAddress, Emailaddresses)" -ForegroundColor Yellow
                Write-Host "6. 'MsolInfo' (Displayname,Firstname,Lastname,Usagelocation,Userprincipalname,signinname,{$_.proxyaddresses,{$_.licenses}) Run Msonline first" -ForegroundColor Yellow
                Write-Host "7. 'Get-MsolAccountSku' List licenses assigned and available, Run Msonline first" -ForegroundColor Yellow
                Write-Host "8. 'Tenantlicense' , Lists used and available licenses(Only available for Parther tenant)" -ForegroundColor White
                Write-Host "9. 'Tenant' (For use in Admin tenant - Lists admin accounts in our Tenants)" -ForegroundColor white
                Write-Host "10. 'HealthStatus' for the active Tenant" -foregroundcolor Yellow
                Write-Host "11. 'Menu' to connect to another tenant/session" -ForegroundColor Yellow
                Write-Host "---------------------------------------------------------------------"
                Write-Host "Example: mailbox | ft" -ForegroundColor cyan
        }
        Function mailbox {
        get-mailbox | Select-Object Displayname, WindowsEmailaddress, PrimarySMTPAddress
                         }
        Function Licensed {
                                                Write-Host "---------------------------------------------------------------------"
                                                Write-Host "Number of Licensed users: "(Get-MsolUser -All | Where-Object {$_.islicensed -like "True"}).count
                                                Write-Host "---------------------------------------------------------------------"
                                                Get-MsolUser -All | Where-Object {$_.islicensed -like "True"}
        }
        Function NotLicensed {
                                                        Write-Host "---------------------------------------------------------------------"
                                                        Write-Host "Number of Unlicensed users: "(Get-MsolUser -All | Where-Object {$_.islicensed -like "false"}).count
                                                        Write-Host "---------------------------------------------------------------------"
                                                        Get-MsolUser -All | Where-Object {$_.islicensed -like "False"}
                                                }     
        Function Mailboxsize {
                $UserMailboxStats = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited @Params | Get-MailboxStatistics
                $UserMailboxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
                $UserMailboxStats | Select-Object DisplayName, TotalItemSizeInBytes,@{Name="TotalItemSize (GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}} | Sort-Object -Property "Totalitemsize (GB)" -Descending
        }
        Function Aliases {
                                                Get-Mailbox | Select-Object Displayname, WindowsEmailAddress, PrimarySmtpAddress, {$_.emailaddresses}
                                        }
        Function MSolInfo {
                                                Get-MsolUser -All | Select-Object Displayname, firstname, lastname,usagelocation, Userprincipalname, signinname,{$_.Licenses.AccountSkuID}, {$_.proxyaddresses}
                                          }
        Function TenantLicense {
                                                                $tenant = Get-MsolPartnerContract
                                                                $tenant | ForEach-Object { $_.name;Get-MsolAccountSku -TenantId $_.tenantid;Write-Host}
                                                                Write-Host "---------------------------------------------------------------------"
                                                        }
        Function Tenant {
                                        $tenant = Get-MsolPartnerContract
                                        foreach ($ten in $tenant) {Get-MsolUser -TenantId $ten.tenantid | Where-Object {$_.userprincipalname -like "admin@*" -and $_.userprincipalname -like "*onmicrosoft.com*"}}
                                        }
        Function Localhost {
                                                        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                                                        Get-PSSession | Remove-PSSession
                                                        $computer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a computer name 'EG: it9-ms-001.it9.local", "Computer", "$env:computername")
                                                        $UserCredential = Get-Credential
                                                        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$computer/PowerShell/ -Authentication Kerberos -Credential $UserCredential
                                                        Import-PSSession $session -AllowClobber
                                                }
        Function O365 {
<#Connect to O365 #>
        <# Remove any preexisting Sessions #>
        Get-PSSession | Remove-PSSession
                                         
        $global:LiveCred = Get-Credential -Credential $global:company
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
        Import-PSSession $Session -AllowClobber
        cls
        tenantmenu
        }
        Function Msonline {
                            <#Connects you to MsOnline services #>
                            Write-Host "-------------------------------------------------------------------------------------------------------------------------------------------------------------"
                            Write-Host "If you wish to manage Microsoft Online Services you will have to download and install the following software, if its  not allready installed on your computer"
                            Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=41950" 
                            Write-Host "And"
                            Write-Host "http://go.microsoft.com/fwlink/p/?linkid=236297"
                            Write-Host "More information:"
                            Write-Host "https://msdn.microsoft.com/en-us/library/azure/jj151815.aspx"
                            Write-Host "-------------------------------------------------------------------------------------------------------------------------------------------------------------"
                            Write-Host ""
                            Write-Host ""


                            Import-Module Msonline
                            Connect-MsolService -Credential $Global:LiveCred
                          }
        Function Sharepoint {
                             #Connecting to SharePoint Online
                             #This first command will import the Sharepoint Online module into your PowerShell session.
                             Write-Host
                             Write-host "---------------------------------------------------------------------------------------------------------------------------------------"
                             Write-Host "To get this function to work, you will have to download and install following software, if its not allready installed on your computer. "
                             Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=35588#"
                             Write-host "More information:"
                             Write-Host "https://support.office.com/en-us/article/Set-up-the-SharePoint-Online-Management-Shell-environment-7b931221-63e2-45cc-9ebc-30e042f17e2c"
                             Write-host "---------------------------------------------------------------------------------------------------------------------------------------"
                             Write-Host ""
                             Write-Host ""
                                
                                Import-Module Microsoft.Online.Sharepoint.PowerShell
                            
                            #Capture administrative credential for future connections.
                            # this is from MS, using allready established instead "$credential = Get-credential"get
                            #Establishes Online Services connection to SharePoint Online
                            
                            #Get the sharepoint URL from your tenant
                            $global:a = $Global:company.Replace("admin@","")
                            $global:b = $global:a.replace(".onmicrosoft.com","")
                            #You must replace the url "https://contoso-admin.sharepoint.com" and use your SharePoint administrative site.
                            Connect-SPOService -url https://$b-admin.sharepoint.com -Credential $Global:LiveCred
 
                         }
        Function Skype4B {
                            Write-host "---------------------------------------------------------------------------------------------------------------------------------------"
                            Write-Host "To get this function to work, you will have to download and install following software, if its not allready installed on your computer."
                            Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=39366"
                            Write-Host "More information:"
                            Write-Host "http://blog.schertz.name/2015/04/managing-lync-online-with-powershell/"
                            Write-host "---------------------------------------------------------------------------------------------------------------------------------------"
                            Write-Host ""
                            Write-host ""
                    
                                Import-Module LyncOnlineConnector
                                $SkypeSession = New-CsOnlineSession -Credential $Global:livecred 
                                Import-PSSession $SkypeSession 
                                Get-CsTenant | fl DisplayName

                         }
        Function HealthStatus {
                                    <# Ignore HTTPS errors #>
                                    [net.servicepointmanager]::ServerCertificateValidationCallback = {$true}
                                    <# Above value can be removed #>
                                    $jsonPayload = (@{userName=$Global:livecred.username;password=$Global:livecred.GetNetworkCredential().password;} | convertto-json).tostring()
                                    $cookie = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/Register" -body $jsonPayload).RegistrationCookie
                                    $jsonPayload = (@{lastCookie=$cookie;locale="en-US";preferredEventTypes=@(0,1)} | convertto-json).tostring()
                                    $events = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/GetEvents" -body $jsonPayload)
                                    $events.events
                              }
        Function Compliance {        
                                $Compliance = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                                Import-PSSession $Compliance -AllowClobber
                            }
<#End functions #>
Menu 