# New-SSLCert
To create a new SSL Cert on Windows with SHA256 or ECC, you have to follow a wizard with lots of steps.

To make this easier, I've created this script so you can create the CSR with just one window.

This is an HTA (HTML Application) file and can be run directly on Windows. There is also a powershell version available that asks the user for the same input.

If you want to run this from the commandline, you can use mshta.exe New-SSLCert.hta

Example screenshot:

![New-SSLCert.hta screenshot](https://github.com/hpmillaard/New-SSLCert/blob/master/New-SSLCert.png?raw=true)

After you've received the certificate from the CA, you can use the import-cer-export-pfx.ps1 script to import the received certificate to the personal store of your computer and export the certificate to a PFX file. You will be asked to point explorer to the cer file, where you want to save the pfx and what password you want to set on the pfx file.
