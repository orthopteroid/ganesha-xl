param([string]$file="")

# https://blogs.u2u.be/u2u/post/creating-a-self-signed-code-signing-certificate-from-powershell

# make new code signing cert and move to root
# $cert = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My -Type CodeSigningCert -Subject "U2U Code Signing" -KeyFriendlyName "U2U" -FriendlyName "U2U"
# Move-Item -Path $cert.PSPath -Destination "Cert:\CurrentUser\Root"

# get cert from root and use to sign file
$cert = Get-ChildItem -Path Cert:\CurrentUser\Root | ? Subject -EQ "CN=U2U Code Signing"
Set-AuthenticodeSignature -FilePath $file -Certificate $cert

# use timestamp to ensure sig is valid after cert expires. abuse? slow?
# https://stackoverflow.com/a/25053511
#Set-AuthenticodeSignature -FilePath $file -Certificate $cert -TimestampServer "http://timestamp.globalsign.com/scripts/timstamp.dll"
