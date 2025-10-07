$certPath = "C:\Path\To\Your\certs\localhost-cert.pem"

# Convert PEM to a cert object
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($certPath)

# Open the Trusted Root Certification Authorities store (Local Machine)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","LocalMachine")
$store.Open("ReadWrite")

# Check if certificate is already installed to avoid duplicates
if (-not ($store.Certificates.Find("FindByThumbprint", $cert.Thumbprint, $false))) {
    $store.Add($cert)
    Write-Host "Certificate imported successfully."
} else {
    Write-Host "Certificate already exists in Trusted Root store."
}

$store.Close()
