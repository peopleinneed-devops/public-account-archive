# to be run in your Exchange server
# inside Exchange powershell module
Set-StrictMode -Version 3.0

# CONFIG. Change to match your env
$MBDname = 'MBD-OFFBOARDING'
$MBDdatadir = "D:\Mailbox\$MBDname"
$MBDlogdir  = "D:\logs\$MBDname"
$MBDserver  = 'mailserver.yourdomain.com'
$ADDC = "dc.yourdomain.com"
# /CONFIG

if (! (Get-Command 'Remove-MailboxDatabase' -ErrorAction SilentlyContinue)) {
    throw "Need to run this from Exchange powershell module"
}



Write-Host "Mailboxes in db $($MBDname): " -NoNewline
$mbxs = Get-Mailbox -Database $MBDname
if ($mbxs) {
    Write-Host $mbxs.count
} else {
    Write-Host 0
}

$mbxs 

Write-Host
Write-Host "Really delete mdb $MBDname a re-create it again?" -ForegroundColor Red
Read-Host "Press [Enter] to continue" | Out-Null

Write-Host "deleting MDB"
Remove-MailboxDatabase -Identity $MBDname -domaincontroller $ADDC 

Write-Host "waiting 10 seconds"
Start-Sleep -Seconds 10

Write-Host "In 'elevated' powershell window do run" -ForegroundColor Cyan
Write-Host "Restart-Service -Name HostControllerService"


Write-Host "ready?"
Read-Host "Press [Enter] to continue" | Out-Null

Write-Host "deleting mdb files on disk"

if (Test-Path $MBDdatadir) {
    Remove-Item -Path $MBDdatadir -Recurse -Force | Out-Null
}

if (Test-Path $MBDlogdir) {
    Remove-Item -Path $MBDlogdir -Recurse -Force | Out-Null
}


write-Host "Now the mdb will be created again"
Read-Host "Press [Enter] to continue" | Out-Null

if (! (Test-Path $MBDdatadir)) {
    New-Item -Path $MBDdatadir -ItemType Directory | Out-Null
}


if (! (Test-Path $MBDlogdir)) {
    New-Item -Path $MBDlogdir -ItemType Directory| Out-Null
}


New-MailboxDatabase -Server $MBDserver -Name $MBDname -domaincontroller $ADDC -EdbFilePath "$MBDDatadir\$($MBDname).edb" -logFolderPath $MBDlogdir 

Write-Host "Added."
Write-Host "Restarting service MSExchangeIS..."

Write-Host "In 'elevated' powershell window do run" -ForegroundColor Cyan
Write-Host "Restart-Service -Name MSExchangeIS"

Read-Host "Press [Enter] to continue" | Out-null

# circular logging - so the logs won't grow up indefinitely
Set-MailboxDatabase $MBDname -CircularLoggingEnabled $true
Dismount-Database $MBDName -Confirm:$false

Mount-Database $MBDname -Confirm:$false
Get-MailboxDatabase $MBDname | Set-MailboxDatabase -IssueWarningQuota "60GB" -ProhibitSendQuota "70GB" -ProhibitSendReceiveQuota "80GB"

Write-Host "Information"
Get-MailboxStatistics -Database $MBDname | ft -AutoSize
Get-MailboxDatabaseCopyStatus * | ft -AutoSize

Write-Host
Write-Host 'Finish.'

