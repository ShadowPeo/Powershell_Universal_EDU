# ==============================================================================================
# NAME: certRenewal.ps1
# DATE  : 21/03/2025
#
# COMMENT: Renews the SSL certificate for the Powershell Universal server using Posh-ACME and Cloudflare
# VERSION: 1, Migration to Powershell Universal from Powershell Core
# ==============================================================================================


# ==============================================================================================
# This script is designed to be run on Powershell Universal not directly as a script and is looking for the following variables to be set in the environment set by the server.
#  - The Script requires the third party module Posh-ACME to be installed and available in the environment
#  - This script requires that the Posh-ACME module has been set up and configured to use the correct servers in the user profile that is being used to run the script
# $certNames - The names of the certificates to be renewed
# $renewDays - The number of days before the certificate expires to renew it
# $cloudflareKey - The API key for Cloudflare (Secret stored as string on the server)
# ==============================================================================================



$certDirectory = "C:\ProgramData\PowerShellUniversal\Certs"

Import-Module Posh-ACME

$pArgs = @{
    CFToken = (ConvertTo-SecureString ($Secret:cloudflareKey) -AsPlainText -Force)
}


#Retrieve Existing Cert if exists
$retrievedCert = Get-PACertificate -MainDomain $certNames[0]

if ($null -eq $retrievedCert) {
    #Request certificate if ite does not exist
    Write-Output "No Certificate Fount Retrieving new Certificate"
    New-PACertificate $certNames -AcceptTOS -Plugin Cloudflare -PluginArgs $pArgs
    
    #Pull newly retrieved certificate
    $retrievedCert = Get-PACertificate -MainDomain $certNames[0]
}
elseif (((New-TimeSpan -Start (Get-Date) -End ($retrievedCert.NotAfter)).Days -le $renewDays)) {
    #Renew cert - add to else if it does exist try renewal
    Write-Output "Certificate due for renewal, ignoring"
    Submit-Renewal -MainDomain $certNames[0]
    $retrievedCert = Get-PACertificate -MainDomain $certNames[0]
}
else {
    Write-Output "Certificate is current, continuing"
}

$serviceRestart = $false

if (Test-Path -Path $certDirectory) {
    #Proceed with Certificate Check
    if (Test-Path -Path (Join-Path -Path $certDirectory -ChildPath "fullchain.cer")) {
        if ((Get-Item $retrievedCert.FullChainFile).LastWriteTime -gt (Get-Item (Join-Path -Path $certDirectory -ChildPath "fullchain.cer")).LastWriteTime) {
            Write-Output "Retrieved Certificate is newer than existing certificate, copying certificate to directory"
            Copy-Item -Path $retrievedCert.FullChainFile -Destination $certDirectory -Force
            $serviceRestart = $true
        }
    }
    else {
        Write-Output "Certificate File not found, exporting retrieved certificate file"
        Copy-Item -Path $retrievedCert.FullChainFile -Destination $certDirectory -Force
        $serviceRestart = $true
    }

    #Proceed With Private Key Test
    if (Test-Path -Path (Join-Path -Path $certDirectory -ChildPath "cert.key")) {
        if ((Get-Item $retrievedCert.KeyFile).LastWriteTime -gt (Get-Item (Join-Path -Path $certDirectory -ChildPath "cert.key")).LastWriteTime) {
            Write-Output "Retrieved Private Key is newer than existing key, copying private key to certificate directory"
            Copy-Item -Path $retrievedCert.KeyFile -Destination $certDirectory -Force
            $serviceRestart = $true
        }
    }
    else {
        Write-Output "Private Key does not exist, copying retrieved private-key to certificate directory"
        Copy-Item -Path $retrievedCert.KeyFile -Destination $certDirectory -Force
        $serviceRestart = $true
    }

    if ($serviceRestart) {
        Write-Output "Changes have been detected, restarting PowershellUniversal Service"
        #Restart the Service/Computer, need to find a way to do this securely without granting admin rights
        #restart-service PowerShellUniversal
    }
}
else {
    Write-Output "Certificate Directory does not exist, please create it and grant the service user permissions to access it"

}