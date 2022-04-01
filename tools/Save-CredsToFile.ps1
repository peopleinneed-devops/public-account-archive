<#
	.SYNOPSIS
		creates .cred file (from your username and password input). Cred file is encrypted via Windows native DPAPI. 
	.DESCRIPTION
		saved credential can be read in plain/used only via the same username that created the file and on the same machine. 
		So it is not transferrable between users or computers.
#>


Write-Host "Insert login (including @yourdomain.com) and press Enter"
$login = Read-Host "Login " 
$login = $login.toLower()

$credFile = "$PSScriptRoot\$login-$($env:username).cred"

Write-Host "Insert password and press Enter"
Write-Host "You can paste it from clipboard on right mouse click"

Read-Host "Password" -AsSecureString | ConvertFrom-SecureString | Out-File $credFile

Write-Host "Saved to file."


<#

# example read
$credFile = (cat $credFile | ConvertTo-SecureString)
(New-Object pscredential "user", $credFile).GetNetworkCredential().Password

#>
