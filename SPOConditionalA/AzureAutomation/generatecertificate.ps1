# Define variables
$certName = "SPOAutomationCert"
$certPassword = "YourStrongPassword123!"  # Change this!
$validityYears = 2
$exportPath = "C:\Certs"

# Create export folder if it doesn't exist
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath
}

# Create self-signed certificate
$cert = New-SelfSignedCertificate -Subject "CN=$certName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($validityYears)

# Get the thumbprint
$thumbprint = $cert.Thumbprint
Write-Host "Certificate Thumbprint: $thumbprint" -ForegroundColor Green

# Export .pfx (Private key - for Azure Automation)
$securePassword = ConvertTo-SecureString -String $certPassword -Force -AsPlainText
Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$thumbprint" `
    -FilePath "$exportPath\$certName.pfx" `
    -Password $securePassword

# Export .cer (Public key - for Azure AD App)
Export-Certificate -Cert "Cert:\CurrentUser\My\$thumbprint" `
    -FilePath "$exportPath\$certName.cer"

Write-Host "`nCertificate files created:" -ForegroundColor Yellow
Write-Host "  PFX (Private): $exportPath\$certName.pfx"
Write-Host "  CER (Public):  $exportPath\$certName.cer"
Write-Host "`nThumbprint: $thumbprint" -ForegroundColor Cyan
Write-Host "Save the thumbprint - you'll need it later!" -ForegroundColor Red