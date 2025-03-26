# Create a self-signed code signing certificate (valid for 1 year)
$cert = New-SelfSignedCertificate -Subject "CN=DoTroubleshooter, O=BAH_ETSS_WinOps" `
    -Type CodeSigningCert `
    -KeyUsage DigitalSignature `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(1) `
    -CertStoreLocation Cert:\CurrentUser\My

# Export the certificate to a PFX file with a password
$password = ConvertTo-SecureString -String "test123" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "$PsScriptroot\CodeSigningCert.pfx" -Password $password
