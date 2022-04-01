#Requires -PSEdition Core
<#
    .SYNOPSIS
        runs whole thing
    .DESCRIPTION
        the scriptname does not contain approved verb as this is not commandlet
#>
Set-StrictMode -Version 3.0

Import-Module ActiveDirectory -ErrorAction Stop -Global



# shared
$includedLibraries = @('\Config.ps1', '\Classes.ps1',  '\Functions.ps1')
foreach ($includedLibrary in $includedLibraries) {
   try {        
       $includedLibrary = $PSScriptRoot + $includedLibrary
       Write-Debug "loading included file $includedLibrary"
       . ($includedLibrary)
   } catch {
       $Error[0]
       Write-Error "FATAL: error loading included file. Either has syntax problem, or does not exist $includedLibrary"    
       exit 1
   }
}

try {
    Stop-Transcript | Out-Null
} catch [System.InvalidOperationException] { }

$TranscriptLog = "$($global:archiveScriptLogDir)$($MyInvocation.MyCommand.Name).transcript.log"
Start-Transcript $TranscriptLog -Append | Out-null


$PSDefaultParameterValues.Clear()
$PSDefaultParameterValues.Add('Write-Log:Path',"$($global:archiveScriptLogDir)$($MyInvocation.MyCommand.Name).log")




Invoke-BeginBlock
$DebugPreference = 'Continue' # zakomentuj pro skryti debug hlasek
$VerbosePreference = 'SilentlyContinue' # zakomentuj pro skryti verbose hlasek

$q = [Queue]::new()
#$q.ADfilterExpression = { (extensionAttribute10 -like 'offboarding') -and (enabled -eq 'false')}
#$q.ADfilterExpression = { (samaccountname -eq 'joukat01') -and (enabled -eq 'false')}
$q.ADfilterExpression = $global:archiveUsersADFilterExpr
$q.ADsearchBase       = 'OU=Uzivatele PIN,DC=pinf,DC=cz'
$q.ADsearchBase       = $global:archiveUsersADSearchBase
$q.limits.[jobPhase]::MoveOnprem     = 20
$q.limits.[jobPhase]::ExportPST      = 20
$q.processUsers()

Invoke-EndBlock