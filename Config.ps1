<#
    .SYNOPSIS
        configuration
#>


######################################################################
## CONFIG ###########################################################
######################################################################

# archive target. Path must be accessible read/write from username running this script AND also by Exchange server itself (add permissions for yourdomain.com\Exchange Trusted Subsystem)
$global:archiveTarget            = '\\san\storage\backups' # without trailing backslash
$global:archiveScriptLogDir      = "c:\logs\" # with trailing backslash
$global:archivedUserPrefixName   = 'Zz[archive]' # name prefix added to archived users. i.e. 'John Smith' => 'Zz[archive] John Smith'. In search, archived users will be displayed last
$global:archivedUsersOU          = "OU=archive,OU=Users,DC=yourdomain,DC=com" # OU path to archived users. Must be existing one
$global:archiveUsersADSearchBase = 'OU=Users,DC=yourdomain,DC=com' # canonical name of OU in Active Directory, that we will search users for
$global:archiveUsersADFilterExpr = { (extensionAttribute10 -like 'offboarding') -and (enabled -eq 'false') } # filter expression for Get-ADUser -Filter, that will return users flagged to be archived

$global:onpremDomainName         = 'yourdomain.com' # your domain name
$global:onpremDomainADDC         = 'dc.yourdomain.com' # fqdn of your domain controller

$global:onpremExchangeMailboxDB  = "MBD-OFFBOARDING" # existing mailbox database (preferably dedicated for archiving process)
$global:onpremExchangeHostname   = 'yourmailserver.yourdomain.com' # fqdn of your onprem exchange
$global:onpremExchangeLogin      = "yourServiceAccount@yourdomain.com" # should have appropriate roles/rights. i.e. be in "Recipient management" 
$global:onpremExchangeCredFile   = "$onpremExchangeLogin-$($env:username).cred"
$global:onpremExchangeCreds      = new-object -typename System.Management.Automation.PSCredential -argumentlist $onpremExchangeLogin,$(cat $onpremExchangeCredFile | convertto-securestring) -ErrorAction stop
$global:onpremExchangeRemoteRA   = 'yourorgname.mail.onmicrosoft.com' # remote routing address to online exchange. should be like [yourOrgname].mail.onmicrosoft.com

# create certificate via
# tools\Create-CertificateForExchangeOnlineAccess.ps1 
# https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
# used method is "connect via certificate thumbprint"
$global:onlineExchangeAppId      = 'xxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'    # appId                               
$global:onlineExchangeCertThumb  = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' # thumprint
$global:onlineExchangeOrg        = 'yourOrgname.onmicrosoft.com' # the default domain for M365. Should be like [yourOrgname].onmicrosoft.com. Check in O365 admin > Settings > Domains

######################################################################
## /CONFIG ###########################################################
######################################################################
