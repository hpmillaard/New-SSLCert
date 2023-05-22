[CmdLetBinding()]
Param(
	[Parameter(Position=0)][string]$CERFile,
	[Parameter(Position=1)][string]$PFXFile,
	[Parameter(Position=2)][string]$PFXPassword
)

If (!$CERFile){
	Add-Type -AN System.Windows.Forms
	Write-Host -f Yellow "CER or CRT File path not provided. Ask user for CER or CRT File"
	$OFD = New-Object System.Windows.Forms.OpenFileDialog
	$OFD.filter = "cer Files (*.cer)| *.cer|crt Files (*.crt)| *.crt"
	$OFD.ShowDialog() | Out-Null
	$CERFile = $OFD.FileName
}
If (!$CERFile){Error "This script can't run without a CER or CRT file and will now quit."}

$cert = Import-Certificate -FilePath $CERFile -CertStoreLocation CERT:\LocalMachine\My\

$repair = certutil -f -repairstore my $cert.thumbprint

If (!$PFXFile){
	Add-Type -AN System.Windows.Forms
	Write-Host -f Yellow "PFX File path not provided. Ask user for PFX File"
	$OFD = New-Object System.Windows.Forms.SaveFileDialog
	$OFD.filter = "PFX Files (*.pfx)| *.pfx"
	$OFD.ShowDialog() | Out-Null
	$PFXFile = $OFD.FileName
}
If (!$PFXFile){Error "This script can't run without a PFX file and will now quit."}

If (!$PFXPassword){Add-Type -AN Microsoft.VisualBasic;$PFXPassword = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Password for the PFX File.','PFX password','P@ssw0rd')}

$pwd = ConvertTo-SecureString -String $PFXPassword -Force -AsPlainText
$cert | Export-PfxCertificate -FilePath $PFXFile -Password $pwd -CryptoAlgorithmOption AES256_SHA256 -Force
pause