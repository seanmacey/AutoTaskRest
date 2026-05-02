# Certiifcate management

## create self signed certificate

```PowerShell
$Cert = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject "CN=Sean Macey" `
    -FriendlyName "Sean Macey PowerShell Code Signing" `
    -CertStoreLocation Cert:\CurrentUser\My `
    -KeyUsage DigitalSignature `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(10)

```

## Export the Public part of the Code Signing Certificate

Use this to get the key to install on Computers that will use the script.

* run this script from the same directory as the AutoTaskRest modules, that was the public certificate will be obvious to

``` Powershell

Export-Certificate `
  -Cert $Cert `
  -FilePath "SeanMacey-CodeSigning-public.cer"

```

### To export both the Public and Private part of the keys togather

Do this to enable you to store the full keychain externally, or to allow you to sign files with the same key from a different computer

* using a password prevents people who gain access to this key from using it unless they have the password
* do this from the cd "Certificates-forAutotask" sub directory

```powershell
$Password = Read-Host "Enter PFX password" -AsSecureString

Export-PfxCertificate `
  -Cert $Cert `
  -FilePath "SeanMacey-CodeSigning-Private-With-Password.pfx" `
  -Password $Password

```

## To Retrieve Self Signing Private Certificate from store

You can then use the key ($Cert) to sign powershell scripts and modules

``` powershell
$Cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object Subject -eq "CN=Sean Macey"

```

### review certificate details

``` powershell
$Cert | Format-List Subject, FriendlyName, Thumbprint, NotAfter

```

## Import the private/Public Certificate Elsewhere

installs the certificate along with private key - so you can sign scripts from that computer also.

```powershell
$Password = Read-Host "Enter PFX 'private {see keeper autotask}' password" -AsSecureString
Import-PfxCertificate `
  -FilePath "SeanMacey-CodeSigning-Private-With-Password.pfx" `
  -CertStoreLocation Cert:\CurrentUser\My `
  -Password $Password

```

## Use the private Cert to Re‑Sign autotaskRes.ps1

* first retrieve the the private $cert, then apply to AutotaskRest.ps1

```powershell
$CerttoApply = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object Subject -eq "CN=Sean Macey"

#Get-ChildItem  -Recurse -Include *.ps1,*.psm1,*.psd1 |
Get-ChildItem   *.ps1,*.psm1,*.psd1 |
  ForEach-Object {
    Set-AuthenticodeSignature $_ $CerttoApply
  }

```

## Trust the public Certificate (execution machine)

run this to install the public certiicate on a machine that will run the autotaskrest module

```powershell
Import-Certificate `
  -FilePath .\SeanMacey-CodeSigning-public.cer `
  -CertStoreLocation Cert:\CurrentUser\TrustedPublisher

Import-Certificate `
  -FilePath .\SeanMacey-CodeSigning-public.cer `
  -CertStoreLocation Cert:\CurrentUser\Root

```
