# About

**Archives (exports) Exchange online mailbox to PST file automatically. No GUI clicking or expensive software needed.**

Tested on hundreds of mailboxes of various size, designed to run multiple exports simultaneously. 
After the successful process the mailbox will exist only in form of PST file.
Common usecase "employee left, now we archive the mailbox for compliance reasons and free the Office365 license".

![Schema](assets/Account-Archive-simple.png?raw=true "Schema")

## Screenshot
![screenshot](assets/screenshot.PNG?raw=true "screenshot")


# Prerequisities
- Exchange online mailboxes
- on-premises (inhouse) Exchange server, cooperating with online Exchange in so called hybrid mode
- a place to store mailbox exports (PST files)

# Installation
- download the code and put its structure into some directory, preferably on server/workstation that can see Onpremise and also Online exchange and has access to some backup folder, where you will archive mailbox exports (PST files). 
- install Powershell Core, the next (and current) generation of Powershell. https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows


## Exchange online unattend connection 
You'll need to set up *app only authentication* via *certificate thumbprint* to Exchange Online. 
By doing so you will benefit from ability to run the archiving process unattended (in Scheduled tasks daily, perhaps) and also having setted up authentication method, that is not going to be deprecated soon (in comparison to login+password authentication metods).

The setup process is described here
https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication
To help with self-signed certificate creation there is an script `tools\Create-CertificateForExchangeOnlineAccess.ps1`

## Exchange onpremises unattend connection
Connection authenticates via saved credentials in encrypted file. Use `tools\Save-CredsToFile.ps1` to save the login/password for the connection. 
Credential file content is protected via native Windows DPAPI. The credential cannot be read by other user or on other machine. So you must save them under the same user and on the same machine, as you will be running the archiving script.

## Exchange onpremises setup
You preferably set-up new mailbox database to be used just for this process, as the database will grow by every new export and it might be easier to re-create it after some months/years than maintain. The mailbox database should have circular logging enabled and enough high quotas to allow ingestion of generously sized online mailboxes. Or you can edit and use `Create-MailboxDatabaseForExchangeOnpremises.ps1` for this task

## configuration for your environment
After you finished required steps, you feed the data into `Config.ps1`. All variables there needs to be filled by correct values for your environment.
It is expected that users ready for archiving process are *marked* in Active directory somehow. The default expectation is, that user's `extensionAttribute10` has value `offboarding` and user account is disabled. For us there are flags coming from our other automation. For you, all of these you can change for your needs. To just test the process and flag some your user as *ready for archivation*, so the account will be processed for archiving, you can issue commands
```
# replace 'someuser' with some actual username
Set-ADUser -Identity someuser -Replace @{'extensionAttribute10'='offboarding'}
Disable-ADAccount -Identity someuser 
```

## additional customizations
at the nearly end of archiving process, after user's mailbox is exported to PST, the phase **CleanAndFinish** begins. 
In this phase it is changing the name of user (with prefix `Zzz[archive]`), so in Active directory search, the user will be intentionally displayed at the very end of results, behind active users. Also other archival processes happening to the AD user account in this phase. 
You might want to ommit them for your environment, change them or perhaps add even a removal of AD account.
You can control that by changing `cleanAndFinish()` method in class `User`.

Also you can limit job concurrency. In default there can be several move jobs running and several PST exports running. You can set it up in `Account-Archive.ps1` via
```
$q.limits.[jobPhase]::MoveOnprem     = 20
$q.limits.[jobPhase]::ExportPST      = 20
```

# Running
open the `pwsh.exe` in the code directory and type there `.\Account-Archive.ps1` followed by Enter. 


