<#
    .SYNOPSIS
        classes User and Queue
        
    .DESCRIPTION
        decision to store all in one .ps1 file is to allow current IDE (Visual Studio Code, as of 2022-03-31)
        to show syntax errors correctly. If splitted into multiple files, VSC would complain about missing class defs etc.

        Object oriented (class) model was preferred for this job's usecase

        The main benefit is that we do not have to pass ADUser and other variables back and forth in cmdlet calls
        and we don't have to validate the input of each cmdlet again

    .EXAMPLE 
        # but the two classes won't work just alone
        $q = [Queue]::new()
        $q.processUsers()
        # which then calls something like
        $u = [User]::new((Get-ADuser -identity $someidentity -properties $properties)) 
        $u.processPhase()
        # the user's archive process state will then be determined and archive process will be started/resumed

#>


# phase names. changing this might break quite a lots of things
enum jobPhase
{    
    Pick
    MoveOnprem
    ExportPST
    CleanAndFinish
    Done
}




class User {
    $ADuser # specification of type [Microsoft.ActiveDirectory.Management.ADAccount] causes in PSCore 7.2.2, that runtime cannot find this type and ends in error.  Alternative solution via 'using module ActiveDirectory' caused even more weird things, so I ommited type specification here
    [hashtable] $mailboxInfo       = [Ordered] @{'onpremType' = $null; 'onlineType' = $null; 'onlineSize' = $null }
    [jobPhase] $currentPhase       = [jobPhase]::Pick # default
    [hashtable] $newJobsAllowances
    [bool] $newJobStarted          = $false

    # constructor
    User($ADuser) {
        $this.init($ADuser)
    }

    # initialization
    [void] init($ADuser)
    {
        <#if (-not ('[Microsoft.ActiveDirectory.Management.ADAccount]' -as [type])) {
            Write-Log -Level 'Error' 'Typ [Microsoft.ActiveDirectory.Management.ADAccount] neni definovan, '
        } else {
            if (Invoke-Expression "$ADuser -isnot [Microsoft.ActiveDirectory.Management.ADAccount]") {
                Write-Log -Level 'Error' 'Ocekavam na vstupu typ [Microsoft.ActiveDirectory.Management.ADAccount], treba z Get-ADUser. Konec'
                exit
            }
        }#>
        $this.ADuser = $ADuser
        if ($null -eq $ADuser.CanonicalName) {
            Write-Error "CanonicalName attribute was expected in ADuser object. But not present. This and maybe more attribtes are missing in input. Exiting" -ErrorAction Stop
            exit 1
        }
        # loads from AD
        $this.loadCurrentJobphase()
        # populate
        $this.mailboxInfo.onpremType = $this.GetMailboxType()
        $this.mailboxInfo.onlineType = $this.GetEOMailboxType()
        $this.mailboxInfo.onlineSize = $this.GetEOMailboxSize()                
    }

    <#
        .SYNOPSIS
            info about detected mailbox attributes (from both onprem and online)
    #>
    [string] getInfo()
    {
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText
        $r = $this.mailboxInfo | Out-String
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::Ansi
        return $r
    }

    <#
        .SYNOPSIS
            alias for logging/output cmdlet, with some enrichment
    #>
    [void] log($message)
    {
        Write-Log -Message ($this.ADuser.SamAccountName + ": " + (($message | out-string).TrimEnd())) 
    }
    [void] log($message, $ForegroundColor='White')
    {
        Write-Log -Message ($this.ADuser.SamAccountName + ": " + (($message | out-string).TrimEnd())) -ForegroundColor $ForegroundColor
    }
    [void] logDebug($message)
    {
        Write-Log -Level Debug -Message ($this.ADuser.SamAccountName + ": " + ($message | out-string))
    }
    [void] logWarning($message)
    {
        Write-Log -Level Warning -Message ($this.ADuser.SamAccountName + ": " + ($message | out-string)) -ForegroundColor Red
    }
    [void] logError($message)
    {
        Write-Log -Level Error -Message ($this.ADuser.SamAccountName + ": " + ($message | out-string))
    }

    <#
        .SYNOPSIS
            loads phase stored in user's attribute in Active directory
    #>
    [void] loadCurrentJobphase()
    {
        try {
            if ($null -eq $this.ADuser.extensionAttribute9) {
                $this.logDebug("phase not found in AD attribute. Setting default phase 'Pick'")
                $this.setPhase([jobPhase]::Pick)
            } elseif ([enum]::IsDefined([jobPhase], $this.ADuser.extensionAttribute9)) {
                $this.currentPhase = [jobPhase]::($this.ADuser.extensionAttribute9)
            } else {
                $this.logWarning("value '$($this.ADuser.extensionAttribute9)' is not found in [jobPhase] set")
            }
        } catch {
            $this.logWarning("cannot load, reason: $($_.ErrorDetails.Message)")
            $this.logError("cannot load, reason: $($_.ErrorDetails.Message)")
        }
    }

    <#
        .SYNOPSIS
            onpremise mailbox type
        .OUTPUTS
            [string] with value if mailbox does not exist
    #>
    [string] GetMailboxType()
    {                   
        if ($null -eq $this.ADuser.mailnickname) {
            $mailboxtype = "DoesNotExist"            
        } else {
            try {
                $mailboxtype = (get-recipient -Identity $this.ADuser.userPrincipalName -ErrorAction Stop).RecipientTypeDetails 
            } catch { 
                $mailboxtype = "DoesNotExist"                
            }
        }        
        return $mailboxtype
    }
    
    <#
        .SYNOPSIS
            online mailbox type
        .OUTPUTS
            [string] with value if mailbox does not exist
    #>
    [string] GetEOMailboxType()
    {
        $user = Get-EOUser -Identity $this.ADuser.UserPrincipalName -ErrorAction SilentlyContinue
        if ($null -eq $user) {
            $mailboxtype = "DoesNotExist"            
        } else {            
            $mailboxtype = $user.RecipientType.ToString()            
        }        
        return $mailboxtype
    }


    <#
        .SYNOPSIS
            online mailbox size
        .OUTPUTS
            [string] with value if mailbox does not exist
    #>
    [string] GetEOMailboxSize() 
    {                        
        $this.logDebug("getting online mbx stats")  

        if ($null -eq $this.ADuser.mailnickname) {
            $this.logDebug("online mailbox missing. user does not have 'mailnickname' AD attribute filled.")
            return 'DoesNotExist'
        }
        try {            
            $mbx = Get-EOMailbox -ResultSize 2 -Identity $this.ADuser.UserPrincipalName -ErrorAction SilentlyContinue
            if ($mbx.MicrosoftOnlineServicesID -ne $this.ADuser.UserPrincipalName) {
                $this.logWarning("for UPN '$($this.ADuser.UserPrincipalName)' was found user '$($mbx.MicrosoftOnlineServicesID). Strange. Online mailbox is considered missing. Please investigate'")
                return 'DoesNotExist'
            } elseif ($mbx.count -ne 1) {
                $this.logWarning("for UPN '$($this.ADuser.UserPrincipalName)' more than one user was found. This should not happen. Exiting")
                exit 1
            }
            
        } catch {            
            $this.logDebug(('online mailbox missing. Get-EOMailbox error message is: ' + $error[0]))
            return 'DoesNotExist'
        }    
        try {
            $UserMailboxStats = Get-EOMailboxStatistics ($this.ADuser.UserPrincipalName) -ErrorAction Stop
        } catch {
            $this.logDebug("online mailbox missing. Get-EOMailboxStatistics error message is: $($error[0]); exception: $($_.ErrorDetails.Message)")  
            return 'DoesNotExist'
        }                    
    
        return (($UserMailboxStats | Select-Object -ExpandProperty TotalItemSize) -replace "(.*\()|,| [a-z]*\)", "")
    }


    <#
        .SYNOPSIS
            checks if online mailbox is expected type
        .OUTPUTS
            [bool] if the mailbox is ready for move
    #>
    [bool] checkEOMailboxType()
    {
        $rcptType = (Get-EOUser -Identity $this.ADuser.UserPrincipalName).RecipientType.ToString()
            $rcptTypeAllowed  = @('UserMailbox')
            $this.logDebug("onlineType = $rcptType")
            $this.logDebug('description: ' + $this.ADuser.description)
            $onpremMbx = (Get-Mailbox -identity $this.ADuser.SamAccountName -ErrorAction SilentlyContinue)
            if ($onpremMbx) {
                $this.log("onprem mailbox exists. here are the data")
                $this.log(($onpremMbx | out-string))
                $this.log("Maybe manually set the phase to 'ExportPST' to fix?")
            }
            if (($rcptType -eq 'MailUser') -and ($this.mailboxInfo.onpremType -eq 'RemoteUserMailbox')) {                
                $this.logWarning("Seems that ONPREM exchange thinks the mailbox is ONLINE whilst ONLINE thinks its ONPREM. Please investigate.")                
                $this.log("If you are ok with loosing onprem mailbox (if its there), run")
                $this.log("Set-ADuser -Clear msExchMailboxGuid,msexchhomeservername,legacyexchangedn,mail,mailnickname,msexchmailboxsecuritydescriptor,msexchpoliciesincluded,msexchrecipientdisplaytype,msexchrecipienttypedetails,msexchumdtmfmap,msexchuseraccountcontrol,msexchversion -Identity " + $this.ADuser.samAccountName)
                $this.log("... wait for AAD sync cycle and then run")
                $this.log("Enable-RemoteMailbox -Identity $($this.ADuser.userPrincipalName) -DomainController $($global:onpremDomainADDC) -Alias $($this.ADuser.samaccountname) -RemoteRoutingAddress $($global:onpremExchangeRemoteRA)") 
            }
            if ($rcptType -notin $rcptTypeAllowed) {
                $this.logWarning("onlineType = '$rcptType' does not match allowed values: [" + ($rcptTypeAllowed -join ',') + ']')                
                return $false
            }
            return $true
    }

    <#
        .SYNOPSIS
            checks if mailbox GUID is of expected type. Checks both online and onprem
        .OUTPUTS
            [bool] if the GUID is correct on both sides
    #>
    [bool] checkExchangeGUID()
    {
        $url = 'https://docs.microsoft.com/en-us/exchange/troubleshoot/move-mailboxes/migrationpermanentexception-when-moving-mailboxes'
         # online guid check
         $onlineExchangeGuid = Get-EOMailbox -Identity $this.ADuser.userPrincipalName -ErrorAction SilentlyContinue | Select-Object ExchangeGuid -ExpandProperty ExchangeGuid  
         # onprem guid check
         $onpremExchangeGuid = Get-RemoteMailbox -Identity $this.ADuser.userPrincipalName -DomainController $global:onpremDomainADDC -ErrorAction SilentlyContinue | Select-Object ExchangeGuid -ExpandProperty ExchangeGuid  

         if (($onlineExchangeGuid -ne $onpremExchangeGuid) -or (! [guid]::TryParse($onpremExchangeGuid, $([ref][guid]::Empty)) )) {
             $this.logWarning("Mailbox ExchangeGuid differs ONPREM/ONLINE ($onpremExchangeGuid / $onlineExchangeGuid). This will prohibit New-EOMoveRequest to work. Try to fix ExchangeGuid manually.")
             $this.log("by")
             $this.log("Set-RemoteMailbox $($this.ADuser.UserPrincipalName) -ExchangeGUID $onlineExchangeGuid -DomainController $global:onpremDomainADDC")
             $this.log("or")
             $this.log("Enable-RemoteMailbox $($this.ADuser.SamAccountName) -Alias $($this.ADuser.SamAccountName) -RemoteRoutingAddress $($this.ADuser.samaccountname)@$($global:onpremExchangeRemoteRA)")
             $this.log("see $url")
             return $false
         }
         return $true
    }

    <#
        .SYNOPSIS
            nazev jobu v onprem a online exchange operacich
    #>
    [string] getJobNameForMailboxOperations()
    {
        return $this.ADuser.SamAccountName + '-archivingtag'
    }

    <#
        .SYNOPSIS
            otestuje pripravenost na danou fazi skriptu
    #>
    [bool] testPreparednessFor([jobPhase] $phase)
    {
        $this.logDebug("phase '$phase' readiness test")

        if ($phase -eq [jobPhase]::Pick) {
            if ($this.ADuser.Enabled -ne $false) {
                $this.logWarning("account is 'enabled', expected was 'disabled'. Only 'disabled' accounts can be processed. Skipping this one")
                return $false
            }              
            if (-not ($this.checkEOMailboxType())) {
                $this.logWarning("please fix manually")                
                return $false
            }  
            if (($this.mailboxInfo.onpremType -eq 'DoesNotExist') -and ($this.mailboxInfo.onlineType -eq 'DoesNotExist')) {
                $this.logWarning("nothing to archive. Mailbox does not exist either onprem or online. Skipping to last phase")   
                $this.setPhase([jobPhase]::CleanAndFinish)
                return $false
            } 
            if (-not $this.checkExchangeGUID()) {
                return $false
            }
        } elseif ($phase -eq [jobPhase]::MoveOnprem) {
           
        } elseif ($phase -in @([jobPhase]::ExportPST, [jobPhase]::CleanAndFinish)) {
            # vytvori uzivatelskou slozku, neni-li
            if (-not (Test-Path -Path $this.getArchivePath())) {
                $this.logDebug(("creating missing user folder in " + $this.getArchivePath()))    
                New-Item -ItemType Directory $this.getArchivePath()
            }
        } 
        return $true
    }

  
    <#
        .SYNOPSIS
            mailbox archive path
        .OUTPUTS
            string (without trailing backslash)
    #>
    [string] getArchivePath()
    {        
        $userFoldername = [string]$this.ADuser.samaccountname + "_" + [string](Convert-DiacriticCharacters($this.ADuser.displayname))
        return "$($global:archiveTarget)\$userFoldername"
    }

    <#
        .SYNOPSIS
            full path to export PST
    #>
    [string] getArchivePathToNewPSTFile()
    {
        $folder = $this.getArchivePath()
        $file   = $this.ADuser.SamAccountName + '__exported_' + (get-date -format "yyyy-MM-dd_HH-mm-ss")  + '.pst'
        return "$folder\$file"              
    }

    <#
        .SYNOPSIS
            various stuff happening to AD account after PST is exported
    #>
    [void] cleanAndFinish()
    {
        $this.saveUserPropertiesToFile()
        $this.removeUserGroupMembership()
        $this.setUserAttributesArchiveFlag()
        $this.moveUserToArchivedOU()
        $this.removeMailbox()                
        # go to final
        $this.setPhase([jobPhase]::Done)
    }

    <#
        .SYNOPSIS
            process the phase
        .OUTPUTS
            [bool] if successful 
    #>
    [bool] processPhase()
    {
        $phase = $this.currentPhase
        $this.logDebug("phase '$phase' starting")
        if (-not ($this.testPreparednessFor($phase))) {
            $this.logWarning("phase '$phase' readiness test FAILED, skipping")            
            return $false
        }

        try {
            if ($phase -eq [jobPhase]::Pick) { 
                $this.setPhase([jobPhase]::MoveOnprem)

            } elseif ($phase -eq [jobPhase]::MoveOnprem) {                                
                $this.moveOnprem()                                

            } elseif ($phase -eq [jobPhase]::ExportPST) {                
                $this.exportPST()

            } elseif ($phase -eq [jobPhase]::CleanAndFinish) {                
               $this.cleanAndFinish()                 

            } elseif ($phase -eq [jobPhase]::Done) {                
                $this.log('[OK] archiving done.', 'green')              

            } else {
                $this.logWarning("Uknown phase '$phase', Did not expect that")
                $this.logError("Uknown phase '$phase', Did not expect that")
                return $false
            }
            return $true

        } catch {
            $this.logWarning("Exception")
            $this.logWarning($_.Exception.ErrorRecord)
            $this.logError("Exception")
            $this.logError($_.Exception.ErrorRecord)
            return $false
        }
    }

    <#
        .SYNOPSIS
            sets the phase to next level and store it to AD user's attribute
    #>
    [void] setPhase([jobPhase] $phase)
    {        
        #$WhatIfPreference = $true
        $this.log("phase '$phase' transition started")
        $this.currentPhase = $phase        
        Set-ADUser -Identity $this.ADuser.SamAccountName -Replace @{'extensionAttribute9' = $phase} -Server $global:onpremDomainADDC
        #$WhatIfPreference = $false
    }



    <#
        .SYNOPSIS
            checks the job to move mailbox onprem or creates new job
    #>
    [bool] moveOnprem()
    {
        $this.log("will check existing job status or create new job to move Online->Onprem")
        $job = Get-EOMoveRequest -BatchName $this.getJobNameForMailboxOperations()
        
        if (-not $job) {
            # we will create new job
            if (-not ($this.newJobsAllowances.($this.currentPhase))) {
                $this.log("Limit for new jobs in this queue exhausted")
                return $false
            }
            try {
                $parameters = @{
                    Identity = $this.ADuser.UserPrincipalName
                    RemoteTargetDatabase = $global:onpremExchangeMailboxDB
                    BatchName = $this.getJobNameForMailboxOperations()
                    RemoteHostName = $global:onpremExchangeHostname
                    RemoteCredential = $global:onpremExchangeCreds
                    TargetDeliveryDomain = $global:onpremDomainName
                    LargeItemLimit = 'Unlimited'
                    BadItemLimit = 'Unlimited'
                    AcceptLargeDataLoss = $true
                    Outbound = $true
                    WarningAction = 'SilentlyContinue'
                } 
                $this.log("creating new EOMoveRequest job")
                $job = New-EOMoveRequest @parameters
                if ($null -eq $job) {
                    $this.logWarning("unable to create the job")
                    return $false
                }
                $this.newJobStarted = $true               
            } catch {
                $this.log("EXCEPTION")
                $this.logWarning($_.ErrorDetails.Message)
                return $false
            }            
            # OK
             if ($job.Status.toString() -eq 'Queued') {
                $this.log("job status: " + $job.Status.toString())
            # problem
            } else {
                $this.log("ERROR")
                $this.logWarning("job status: " + $job.Status.toString())
            }
        } else {
            # job exists, check state
            $jobStats = $job | Get-EOMoveRequestStatistics
            if ($job.Status.ToString() -like 'Completed*') { 
                # also matches: Completed, CompletedWithWarnings
                $this.log("job status: " + $job.Status.ToString())                
                Remove-EOMoveRequest -Identity $this.ADuser.userPrincipalName -Force -Confirm:$false -Erroraction Stop # via $job and pipe is the send to remove-Eomoverequest not very reliable
                $this.setPhase([jobPhase]::ExportPST)
            } elseif ($job.Status.ToString() -in @('Failed', 'Suspended', 'AutoSuspended')) {                
                $this.logWarning("job status: " + $job.Status.ToString() + ", percent complete: " + $jobStats.PercentComplete + ", message: " + $jobStats.Message + " (failureType: " + $jobStats.FailureType + ")") 
                if ($job.Status.ToString() -eq 'Failed') {
                    # give it a try again                    
                    $this.log('resuming (to try again)...')
                    # perhaps a bug in Resume command that -errorVariable nor -warningVariable parameters does not store the error
                    # neither try/catch block catches it. But error is stored in global $error variable. So this workaround
                    $lastError = $error[0]
                    $job | Resume-EOMoveRequest
                    if ($lastError -ne $error[0]) {                        
                        $this.logWarning("$($error[0])")
                        $this.logWarning('job NOT resumed, try to fix manually')                        
                    } else {
                        $this.log('job resumed')
                    }                    
                }
            } else {                
                $this.log("job status: " + $job.Status.ToString() + ", percent complete: " + $jobStats.PercentComplete, 'blue') 
            }

        }
        return $true
    }

    <#
        .SYNOPSIS
            checks the job to export to PST or creates new job
    #>
    [bool] exportPST()
    {
        $this.log("will check existing job status or create new job for export to PST")
        $job = Get-MailboxExportRequest -BatchName $this.getJobNameForMailboxOperations()
        
        if (-not $job) {
            # make a new one   
            if (-not ($this.newJobsAllowances.($this.currentPhase))) {
                $this.log( "Limit for new jobs in this queue exhausted")
                return $false
            }       
            try {
                $parameters = @{
                    Mailbox = $this.ADuser.SamAccountName                    
                    BatchName = $this.getJobNameForMailboxOperations()                    
                    DomainController = $global:onpremDomainADDC                  
                    BadItemLimit = 'Unlimited'
                    LargeItemLimit = 'Unlimited'
                    AcceptLargeDataLoss = $true
                    FilePath     = ($this.getArchivePathToNewPSTFile())
                    WarningAction = 'SilentlyContinue'
                } 
                $this.logDebug("creating new MailboxExportRequest job")
                $job = New-MailboxExportRequest @parameters
                $this.newJobStarted = $true

                # OK
                if ($job.Status.toString() -eq 'Queued') {
                    $this.log("job status: " + $job.Status.toString())
                # problem
                } else {
                    $this.log("ERROR")
                    $this.logWarning("job status: " + $job.Status.toString())
                }
            } catch {
                $this.log("EXCEPTION")
                $this.logWarning($_.ErrorDetails.Message)
            }            
        } else {
            # check the state of existing
            $jobStats = $job | Get-MailboxExportRequestStatistics
            if ($job.Status.ToString() -like 'Completed*') { 
                # also matches: Completed, CompletedWithWarnings
                $this.log("job status: " + $job.Status.ToString())
                $job | Remove-MailboxExportRequest -Confirm:$false -Erroraction Stop
                
                $this.setPhase([jobPhase]::CleanAndFinish)
            } elseif ($job.Status.ToString() -in @('Failed', 'Suspended', 'AutoSuspended')) {                
                $this.logWarning("job status: " + $job.Status.ToString() + ", percent complete: " + $jobStats.PercentComplete + ", message: " + $jobStats.Message + " (failureType: " + $jobStats.FailureType + ")") 
                if ($job.Status.ToString() -eq 'Failed') {
                    # give it a try again                    
                    $this.log('resuming (to try again)...')
                    # perhaps a bug in Resume command that -errorVariable nor -warningVariable parameters does not store the error
                    # neither try/catch block catches it. But error is stored in global $error variable. So this workaround
                    $lastError = $error[0]
                    $job | Resume-MailboxExportRequest 
                    if ($lastError -ne $error[0]) {
                        $this.logWarning("$($error[0])")
                        $this.logWarning('job NOT resumed, try to fix manually')
                    } else {
                        $this.log('job resumed')
                    }
                }
            } else {                
                $this.log("job status: " + $job.Status.ToString() + ", percent complete: " + $jobStats.PercentComplete, 'blue') 
            }

        }   
        return $true     
    }

    <#
        .SYNOPSIS
            changes user's AD attributes (title, name, etc)
    #>
    [void] setUserAttributesArchiveFlag()
    {         
        $newTitle = '[archived] not an employee'
        $newDesc  = "archived " + (get-date -format "yyyy-MM-dd HH:mm:ss") + " by script " + (split-path $pscommandpath -leaf)

        $this.logDebug("Clearing EA10, setting job title='$newTitle' and description='$newDesc'")
                
        Set-ADuser -identity $this.ADuser.SamAccountName -clear extensionAttribute10 -Server $global:onpremDomainADDC -Erroraction Stop
        Set-ADUser -identity $this.ADuser.SamAccountName -title $newTitle -Server $global:onpremDomainADDC -Erroraction Stop
        Set-ADUser -identity $this.ADuser.SamAccountName -description $newdesc -Server $global:onpremDomainADDC -Erroraction Stop

        # archive name prefix
        $this.logDebug("renaming username to archive username")
        $a = Get-AdUser -identity $this.ADuser.SamAccountName -server $global:onpremDomainADDC
        $a | Rename-ADObject -NewName ($global:ArchivedUserPrefixName + ' ' + $a.Name) -Server $global:onpremDomainADDC
    }


    <#
        .SYNOPSIS
            moves user AD object to specific Organizational Unit
    #>
    [void] moveUserToArchivedOU()
    {
        $this.logDebug("moving to archive OU")
        try {                                
            # odstraneni ochrany            
            Get-AdUser -identity $this.ADuser.SamAccountName -Server $global:onpremDomainADDC | Set-ADObject -ProtectedFromAccidentalDeletion $false -Server $global:onpremDomainADDC            
            # presun
            Get-AdUser -identity $this.ADuser.SamAccountName -Server $global:onpremDomainADDC | Move-ADObject -Targetpath $global:ArchivedUsersOU -Confirm:$false -Server $global:onpremDomainADDC
        } catch {
            Write-Log -Info "Error: $($this.ADuser.SamAccountName) - Move to archived OU failed,  $($_.ErrorDetails.Message)"            
        }
    }

    <#
        .SYNOPSIS
            removes mailbox. no longer needed after pst export done
    #>
    [void] removeMailbox()
    {        
        if ($this.mailboxInfo.onpremType -ne 'DoesNotExist') {
            $this.logDebug("removing mailbox")
            Disable-Mailbox -Identity $this.ADuser.SamAccountName -IgnoreLegalHold -ErrorAction Stop -Confirm:$false -DomainController $global:onpremDomainADDC
        } else {            
            $this.logWarning("mailbox does not exist, so we cannot remove it")
            $this.logDebug("removing selected extensionAttributes instead (EA8,EA10)")
            Set-ADUser -Identity $this.ADuser.SamAccountName -Clear extensionAttribute9,extensionAttribute10 -Server $global:onpremDomainADDC
        }
    }

    <#  
        .SYNOPSIS
            stores all ADuser attributes to file
        .OUTPUTS
            file at $archiveTarget\$samAccountName.txt
    #>
    [void] saveUserPropertiesToFile()
    {
        $this.logDebug("backup of present ADuser attributes to file")
        $filePath = ($this.getArchivePath() + "\" + $this.ADuser.SamAccountName + '.txt')
        # causes, that out-file will not store colored console output (ascii control characters). 
        # i.e. instead
        #[32;1mAccountExpirationDate                 : [0m
        # will be in file
        # AccountExpirationDate                 : 0
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText
        (get-date -format "yyyy-MM-dd HH:mm:ss") + "`r`n`r`n" | Out-File -Encoding utf8NoBOM -Append $filePath
        Get-ADUser -Identity $this.ADuser.SamAccountName -Properties * -Server $global:onpremDomainADDC | Out-File -Encoding utf8 -Append $filePath
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::Ansi
    }


    <#
        .SYNOPSIS
            removes all user group membership (except Domain userss)
    #>
    [void] removeUserGroupMembership()
    {
        $this.logDebug("removing group membership")
        $groups = Get-AdPrincipalGroupMemberShip $this.ADuser.SamAccountName -Server $global:onpremDomainADDC -resourceContextServer $global:onpremDomainName | Where-Object {$_.Name -ne 'Domain Users'} 

        foreach ($group in $groups) {
            $groupname = [string]$group.distinguishedName
            try {               
                Remove-ADPrincipalGroupMembership -Identity $this.ADuser.SamAccountName -memberOf $groupname -Server $global:onpremDomainADDC -Confirm:$false
            } catch {
                $this.logWarning("$($this.ADuser.SamAccountName) - failed to remove AD group membership: '$groupname'; reason: $($_.ErrorDetails.Message)")
            }            
        }
    }

} # class





class Queue {
    # how many parallel jobs can run in each phase (if phase not named = unlimited)
    $limits = @{ 
        [jobPhase]::MoveOnprem     = 5
        [jobPhase]::ExportPST      = 5
    }
    # actual jobs running in each phase. will be filled on init
    $actualRuns = @{ 
        [jobPhase]::Pick           = $null
        [jobPhase]::MoveOnprem     = $null
        [jobPhase]::ExportPST      = $null
        [jobPhase]::CleanAndFinish = $null
        [jobPhase]::Done           = $null
    }
    # where to search in AD
    $ADsearchBase                  = $null 
    # filter to use in AD search
    $ADfilterExpression            = $null
    
    Queue()
    {
        $this.populate()
    }

    <#
        .SYNOPSIS
            detect current queue state
    #>
    [void] populate()
    {
        Write-Log -Level Debug "Queue: populate fronty"
        $this.actualRuns.[jobPhase]::Pick           = 0
        $this.actualRuns.[jobPhase]::MoveOnprem     = 0
        $this.actualRuns.[jobPhase]::ExportPST      = 0
        $this.actualRuns.[jobPhase]::CleanAndFinish = 0
        $this.actualRuns.[jobPhase]::Done           = 0   
        try {
            $this.actualRuns.[jobPhase]::MoveOnprem = (Get-EOMoveRequest | Measure-Object).count
            $this.actualRuns.[jobPhase]::ExportPST  = (Get-MailboxExportRequest | Measure-Object).count
        } catch {
            Write-Log $_.Exception.Message
        }         
    }

    <#
        .OUTPUTS
            number of actual running jobs
    #>
    [int] getActualRuns([jobPhase] $phase)
    {                
        return $this.actualRuns.$phase
    }

    <#
        .OUTPUTS
            current job limit or string 'Unlimited' 
    #>
    [string] getLimit([jobPhase] $phase)
    {
        if ($this.limits.ContainsKey($phase)) {
            return $this.limits.$phase
        }
        return 'Unlimited'
    }

    <#
        .OUTPUTS
            $true if can start new job
    #>
    [bool] canStartNewJob([jobPhase] $phase)
    {
        if (-not ($this.limits.ContainsKey($phase))) {
            # limit neni stanoven
            return $true
        }
        return ($this.actualRuns.$phase -lt $this.limits.$phase)
    }

    <#
        .SYNOPSIS
            used by User object (it is passed to him), so he knows wheter to start new job or not
        .OUTPUTS
            hashtable [phasename->boolean] of phases and their allowances
    #>
    [hashtable] getAllowedNewJobsByPhase() {
        $r = @{}
        $this.actualRuns.GetEnumerator() | foreach-object {
            $r.Add($_.key, $this.canStartNewJob($_.key))
        }
        return $r
    }
    

    <#
        .SYNOPSIS
            manager of queue
        .DESCRIPTION
            creates new user objects and manages their run
    #>
    [void] processUsers()
    {
        Write-Log -ForegroundColor Cyan -Message ([string]((Get-PSCallStack)[0].FunctionName) + " started")

        if (($null -eq $this.ADfilterExpression) -or ($null -eq $this.ADsearchBase)) {
            Write-Log -Level Error "ADfilterExpression or ADsearchBase is null. These must be set to correct value. Exiting"
            exit 1        
        }

        Write-Log -Level Debug "filterExpression: $($this.ADfilterExpression)"        
        Write-Log -Level Debug "searchBase: $($this.ADsearchBase)"

        $properties = @('samaccountname','title','name', 'displayname', 'CanonicalName', 'userPrincipalName', 'extensionAttribute1', 'department','enabled','mail', 'mailnickname', 'whenCreated', 
                'whenChanged', 'lastLogonDate', 'description', 'msExchRemoteRecipientType', 'msExchRecipientTypeDetails', 'extensionattribute9', 'extensionattribute10')
        # order of sort. The first mentioned phase will be on the bottom of results (resulting it will be processed lastly), and so on
        $sortImportanceFromLast = @( '', 'Pick', 'Done', 'CleanAndFinish', 'MoveOnprem', 'ExportPST')
        try {
            $filterResult = Get-ADUser -Filter $this.ADfilterExpression -searchBase $this.ADsearchBase -Properties $properties | Sort-Object { $sortImportanceFromLast.IndexOf($_.extensionAttribute9) } -Descending   
        } catch {
            throw $_.ErrorDetails.Message
            exit 1
        }
        
        Write-Log ("Summary: " + (($filterResult | Measure-Object).Count) + " users in a Queue to be archived")
        Write-Log -Level Debug "List of them:"
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText        
        Write-Log -Level Debug (($filterResult | Select-Object samaccountname, name, extensionAttribute9 | out-string))
        $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::Ansi
        Write-Log -Level Debug "-------------- ----"
        
        
        $filterResult | foreach-object {    

            # User objekt
            Write-Log -Level Debug "$($_.samAccountName): about to load the user..."        
            $u = [User]::new($_)                        
            Write-Log "$($u.ADUser.samAccountName): loaded user data in phase '$($u.currentPhase)'"
            Write-Log -Level Debug $u.getInfo()          

            do {
                $previousPhase = $u.currentPhase
                
                Write-Log -Level Debug ("Queue '$($u.currentPhase)' jobs running/limit: " + $this.getActualRuns($u.currentPhase) + " / " + $this.getLimit($u.currentPhase))
                                
                # je ve fronte misto?
                $u.newJobsAllowances = $this.getAllowedNewJobsByPhase()
                $u.processPhase()    
                
                if ($u.newJobStarted) {
                    $this.populate()
                }
                    
                
            } while ($u.currentPhase -ne $previousPhase) # dokud se to nejak posouva

            Write-Log  "$($u.ADUser.samAccountName): leaving user, in phase '$($u.currentPhase)'"
        }
        Write-Log "all users processed"
        Write-Log ('-' * 40)
        Write-Log ([string] $this.actualRuns.[jobPhase]::MoveOnprem + " `tjobs MoveOnprem")
        Write-Log ([string] $this.actualRuns.[jobPhase]::ExportPST  + " `tjobs ExportPST")
        Write-Log ('-' * 40)
    }
} # class Queue


