# Sign-Output.ps1
# Signs all exe, dll and msix and ps1 Files in a folder including Subfolders)
# Parameters:
#    -Path: Folder path of the files that are to be signed
#    -CertificateThumbprint: Thumbprint of signing certificate to use (have to be installed in personal certificate store)
#    -SigntoolPath: Path of Microsoft signtool (default c:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe)
#    -SelfSignedName: CN used for SelfSigning

param(
    [Parameter(        
        Mandatory=$false,
        HelpMessage="Sign the files in the given folder with a signing certificate"
     )]
     [string]$Path="",
     [string]$CertificateThumbprint="",
     [string]$SigntoolPath="C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x86\signtool.exe",
     [string]$SelfSignedName="Contoso_CustomShell"
)

if ("" -eq $Path)
{
    Write-Host "no path parameter given, exiting"
    Write-Host "Usage:"
    Write-Host "    Sign-Files -path <folder of files to sign> [-CertificateThumbprint <Thumprint of signing certificate to use>]"
    Write-Host "    if -CertificateThumbprint is omitted, a SelfSigned $CertificateThumbprint is used"

    exit
}

# If no thumbprint provided, try to find or create our own SelfSigned cert
if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {

    # Try to find existing certificate
    $existing = Get-ChildItem Cert:\CurrentUser\My |
                Where-Object { $_.Subject -eq "CN=$SelfSignedName" }

    if ($existing) {
        $CertificateThumbprint = $existing[0].Thumbprint
    }
    else {
        # Create new certificate
        $newCert = New-SelfSignedCertificate `
            -Type Custom `
            -Subject "CN=$SelfSignedName" `
            -KeyUsage DigitalSignature `
            -FriendlyName "SelfSignCert" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -TextExtension @(
                "2.5.29.37={text}1.3.6.1.5.5.7.3.3", 
                "2.5.29.19={text}"
            )

        $CertificateThumbprint = $newCert.Thumbprint
    }
}

# Get cert from store
$cert = Get-ChildItem -path Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }

if (-not $cert)
{
    Write-Host "Certificate $CertificateThumbprint not found"
    exit
}

# Get all candidate files (dll, exe, ps1) recursively
$files = Get-ChildItem -Path $Path -Include *.dll,*.exe,*.ps1,*.msix -File -Recurse

foreach ($file in $files)
{
    $LastWrite = $file.LastWriteTime

    # Check current signature
    $currentSig = Get-AuthenticodeSignature -FilePath $file
    $mySubject = $cert.Subject

    if ($currentSig.SignerCertificate) {
        $sigCert = $currentSig.SignerCertificate

        if ($sigCert.Subject -ne $mySubject) {
            # Signed by someone else (expired or valid) -> skip
            continue
        }

        if ($currentSig.Status -eq 'Valid' -and $sigCert.Thumbprint -eq $CertificateThumbprint) {
            # Already signed with current valid cert -> skip
            continue
        }
    }
    Write-Host "$($file.FullName) --> Signing..." -ForegroundColor Green

    try {
        # Signing will fail if the file is compiled as AnyCpu with the flag "Prefer 32-bit"
        # If that happens, remove that flag, or set the build to either x86 or x64
        & "$SigntoolPath" sign /sha1 $CertificateThumbprint /fd SHA256 /td SHA256 /v "$($file.FullName)"

        # Restore LastWriteTime of signed file
        $file.LastWriteTime = $LastWrite
    }
    catch {
        Write-Host "Failed to sign $($file.FullName): $($_.Exception.Message)" -ForegroundColor Red
    }
}
