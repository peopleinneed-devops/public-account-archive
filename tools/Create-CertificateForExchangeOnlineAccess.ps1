<#
    .SYNOPSIS
        creates certificate for Powershell access to ExchangeOnline
    .LINK
        https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate
#>

$certName = "ExchangeOnline - automation on server $($env:COMPUTERNAME) under user $($env:USERNAME)" ## Replace 
$certExportPath = "C:\Users\$($env:USERNAME)\Desktop\$certName"

$cert = New-SelfSignedCertificate -Subject "CN=$certName" -CertStoreLocation "Cert:\CurrentUser\My" `
     -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
     -NotAfter (Get-Date).AddYears(10)

#Export-PfxCertificate -Cert $cert -FilePath ($certExportPath + '.pfx') -ProtectTo ('pinf\' + $($env:USERNAME)) | out-null
Export-Certificate -Cert $cert -FilePath ($certExportPath + '.cer') | out-null

Write-Host "Exported to: $certExportPath"
Write-host 'Done.'