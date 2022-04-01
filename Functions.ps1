<#
    .SYNOPSIS
        various support functions
#>

Function Write-Log {
    <#
	.SYNOPSIS
		Writes submitted text to console and to log
    .EXAMPLE
        $PSDefaultParameterValues.Clear()
        $PSDefaultParameterValues.Add('Write-Log:Path',"c:\AUTOMAT\logs\$($MyInvocation.MyCommand.Name)_" + (Get-Date -format "yyyy-MM-dd__HH-mm-ss") + ".log")
        $DebugPreference = 'continue'
        Write-Log -Level "Debug" "text to log"
	#>

    [cmdletbinding()]
    Param(
        [Parameter(ValueFromPipeline = $True)] $Message,   
        [ValidateSet("Error", "Warning", "Host", "Output", "Verbose", "Info", "Debug")] [string] $Level = "Host",                        
        [String] $ForegroundColor = 'White',        
        [ValidateRange(1, 30)] [Int16] $Indent = 0,                
        [Parameter()] [ValidateScript( {Test-Path $_ -IsValid})] [ValidateScript( {$_ -match '\.\w+'})] # filepath expected
        [string] $Path,
        [Parameter()] [Switch] $OverWrite        
    )

    Begin {
        # if object, then convert
        if ($Message -and $Message.gettype() -ne 'String') {
            $Message = $Message | Out-String
        }        
        if ($Level -eq 'Host') { 
            $Level = 'Info' 
        }
    }

    Process {
            $ErrorActionPreference = 'continue' 
            if ($Message) {                
                switch ($Level) {
                    'Error' { Write-Error $Message }
                    'Warning' {  Write-Host ('WARNING: ' + $Message) -ForegroundColor Red -NoNewline }
                    'Info' { Write-Host $Message -ForegroundColor $ForegroundColor -NoNewline}
                    'Output' { Write-Output $Message }
                    'Verbose' { $message = $message.TrimEnd(); Write-Verbose $Message }
                    'Debug' { $message = $message.TrimEnd(); Write-Debug $Message; }
                }
                $ErrorActionPreference = 'stop' # has to be after Write-Error otherwise this try/catch block will be affected
            }                        
			
            # final text to write to log
            $Message = $Message.TrimEnd()
            if ($Message -and $Message.contains("`n")) {                
                $Message = @($Message -split "`n")
                # first line
                $msg = "{0}{1} [{2}]`t{3}" -f (" " * $Indent), (Get-Date -Format s), $Level.ToUpper(), $Message[0]
                # other lines
                for ($i = 1; $i -lt $Message.count; $i++) {
                    $msg += "`n{0}{1} [{2}]`t{3}" -f (" " * $Indent), (Get-Date -Format s), $Level.ToUpper(), $Message[$i]
                }
            } else {
                # single line message
                $msg = "{0}{1} [{2}]`t{3}" -f (" " * $Indent), (Get-Date -Format s), $Level.ToUpper(), $Message
            }

            $CommandParameters = @{
                FilePath    = $Path
                Encoding    = 'UTF8'
                ErrorAction = 'stop'
                Append      = $true
            }            
            # writing to log            
            $msg | Out-File @CommandParameters
	        
    } #End Process

}




<#
    .SYNOPSIS
        removes diacritics
    .EXAMPLE
        PS C:\> Convert-DiacriticCharacters "Ångström"
        Angstrom
        PS C:\> Convert-DiacriticCharacters "Ó señor"
        O senor
    .OUTPUTS
        [string] ascii
    .LINK 
        https://stackoverflow.com/a/46660695
#>
function Convert-DiacriticCharacters {
    param(
        [string]$inputString
    )
    [string]$formD = $inputString.Normalize(
            [System.text.NormalizationForm]::FormD
    )
    $stringBuilder = new-object System.Text.StringBuilder
    for ($i = 0; $i -lt $formD.Length; $i++){
        $unicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($formD[$i])
        $nonSPacingMark = [System.Globalization.UnicodeCategory]::NonSpacingMark
        if($unicodeCategory -ne $nonSPacingMark){
            $stringBuilder.Append($formD[$i]) | out-null
        }
    }
    $stringBuilder.ToString().Normalize([System.text.NormalizationForm]::FormC)
}



<#
    .SYNOPSIS
        connects to online Exchange
#>
function Connect-ExchangeOnlineCustom($appId, $certThumb, $org) {    
    Import-Module ExchangeOnlineManagement    
    $isConnected = Get-PSSession | where-object { ($_.State -eq 'Opened') -and ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.ComputerName -eq 'outlook.office365.com') }
    if (! $isConnected) {        
        try {
            Write-Log "Connect-ExchangeOnline"
            Connect-ExchangeOnline -Prefix "EO" -appId $appId -certificateThumbprint $certThumb -organization $org -ShowBanner:$false            
        } catch {
            Write-Log "FATAL: Unable to connect" 
            Write-Error $Error[0]
            exit 1
        }
    }
}

<#
    .SYNOPSIS
        connects to onprem exchange
#>
function Connect-ExchangeOnpremCustom {
    try {        
        $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $global:onpremExchangeLogin,$(cat $global:onpremExchangeCredFile | convertto-securestring) -ErrorAction stop
        $isConnected = Get-PSSession | where-object { ($_.State -eq 'Opened') -and `
                ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.ComputerName -eq $global:onpremExchangeHostname) }
        if (! $isConnected) {
            Write-Log "Connect-ExchangeOnprem"
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($global:onpremExchangeHostname)/PowerShell/" -Authentication Kerberos -Credential $Cred -ErrorAction stop
            Import-PSSession $Session -ErrorAction stop | out-null
        }        
    }
    catch {
        Write-Log "FATAL: Unable to connect to On-Prem Exchange"             
        Write-Error $Error[0]   
        exit 1      
    }
}

<#
    .SYNOPSIS
        global init block
#>
function Invoke-BeginBlock
{
    Write-Log ('=' * 80)
    Write-Debug ($((Get-PSCallStack)[0].FunctionName) + " started") 
    #Remove-AllSessions        
    Connect-ExchangeOnlineCustom -appId $global:onlineExchangeAppId -certThumb $global:onlineExchangeCertThumb -org $global:onlineExchangeOrg
    Connect-ExchangeOnpremCustom
    Write-Debug ($((Get-PSCallStack)[0].FunctionName) + " finished") 
}
function Invoke-EndBlock
{
    Write-Debug ($((Get-PSCallStack)[0].FunctionName) + " started") 
    Stop-Transcript | Out-null
    Write-Debug ($((Get-PSCallStack)[0].FunctionName) + " finished") 
}
