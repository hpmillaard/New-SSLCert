If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {Start powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit}
Write-Host -F Green " Ask user input"
Add-Type -AssemblyName Microsoft.VisualBasic
$Name = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Certificate Common Name','Name')
$Country = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Country for the Certificate','Country')
$State = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Province or State for the Certificate','Province or State')
$City = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the City for the Certificate','City')
$Organisation = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Organisation for the Certificate','Organisation')
$OU = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Organisational Unit for the Certificate','Organisational Unit')
$DNS = [Microsoft.VisualBasic.Interaction]::InputBox('Optional: Enter the DNS Alternate names for the certificate. Separate with ;','SAN Names')

$CN = 'CN=' + $Name + ',C=' + $Country + ',S=' + $State + ',L=' + $City + ',O=' + $Organisation + ',OU=' + $OU
Write-Host -f Green " $CN"

Write-Host -f Green " create request.inf"
$F = "$ENV:TEMP\Request.inf"
sc $F '[Version]'
ac $F 'Signature="$Windows NT$"'
ac $F '[NewRequest]'
ac $F "Subject = ""$($CN)"""
ac $F 'Exportable = True'
ac $F 'ProviderName = "Microsoft Software Key Storage Provider"'
ac $F 'KeyLength = 2048'
ac $F 'KeyAlgorithm = RSA'
ac $F 'HashAlgorithm = SHA256'
ac $F 'MachineKeySet = True'
ac $F 'SMIME = False'
ac $F 'UseExistingKeySet = False'
ac $F 'RequestType = PKCS10'
#ac $F 'KeyUsage = 0x1'			#Encipher Only
#ac $F 'KeyUsage = 0x2'			#(Offline) CRL Signing
#ac $F 'KeyUsage = 0x4'			#Key Certificate Signing
#ac $F 'KeyUsage = 0x8'			#Key Agreement
#ac $F 'KeyUsage = 0x10'		#Data Encipherment
#ac $F 'KeyUsage = 0x20'		#Key Encipherment
#ac $F 'KeyUsage = 0x40'		#Non Repudiation
#ac $F 'KeyUsage = 0x80'		#Digital Signature
#ac $F 'KeyUsage = 0x8000'		#Decipher Only
ac $F 'KeyUsage = 0xA0'			#Digital Signature + Key Encipherment
ac $F 'Silent = True'
ac $F '[EnhancedKeyUsageExtension]'
#ac $F 'OID=1.3.6.1.4.1.311.10.3.4'	#Encrypting File System
#ac $F 'OID=1.3.6.1.4.1.311.10.3.12'#Document Signing
#ac $F 'OID=1.3.6.1.4.1.311.20.2.2'	#Smart Card Logon
#ac $F 'OID=1.3.6.1.4.1.311.54.1.2'	#Remote Desktop
#ac $F 'OID=1.3.6.1.4.1.311.80.1'	#Document Encryption
ac $F 'OID=1.3.6.1.5.5.7.3.1'		#Client Authentication
ac $F 'OID=1.3.6.1.5.5.7.3.2'		#Server Authentication
#ac $F 'OID=1.3.6.1.5.5.7.3.3'		#Code Signing
#ac $F 'OID=1.3.6.1.5.5.7.3.4'		#Secure E-mail (S/MIME)
#ac $F 'OID=1.3.6.1.5.5.7.3.5'		#IP Security End System
#ac $F 'OID=1.3.6.1.5.5.7.3.6'		#IP Security Tunnel Endpoint
#ac $F 'OID=1.3.6.1.5.5.7.3.7'		#IP Security User
#ac $F 'OID=1.3.6.1.5.5.7.3.8'		#Time Stamping
#ac $F 'OID=1.3.6.1.5.5.7.3.9'		#OCSP Signing
#ac $F 'OID=1.3.6.1.5.5.7.3.17'		#IP Security Key Exchange (IKE)
#ac $F 'OID=1.3.6.1.5.5.7.3.21'		#Secure Shell Client Authentication
#ac $F 'OID=1.3.6.1.5.5.7.3.22'		#Secure Shell Server Authentication
#ac $F 'OID=2.5.29.37.0'			#any Extended Key Usage
ac $F '[Extensions]'
ac $F '2.5.29.17 = "{text}"'
If ($DNS){
	$DNSA = $DNS.Split(";")
	$DNSA | Foreach {ac $F "_continue_ = ""dns=$($_)&"""}
}Else{
	ac $F "_continue_ = ""dns=$Name&"""
}
	
$CSRName = $Name.replace('*','Wildcard')
$CSR = "$PSScriptRoot\$CSRName.csr"
certreq.exe -new $F $CSR
Write-Host -f Green " $CSR created"
del $F