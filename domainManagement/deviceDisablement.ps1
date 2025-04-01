# ==============================================================================================
# NAME: Device Disablement.ps1
# DATE  : 28/03/2025
#
# COMMENT: Converts EduHub SF table (csv file) into standard ASM import file
# VERSION: 1, Migration to Powershell Universal from Powershell Core
# VERSION: 1.1 - Cleaned up code and changed how the AD changes are made through credentials rather than using the current user
# ==============================================================================================


# ==============================================================================================
# This script is designed to be run on Powershell Universal not directly as a script and is looking for the following variables to be set in the environment set by the server.
#  - As this uses active directory it is assumed that the server is a member of the domain and has access to the Active Directory servers. and the Active Directory module is installed.
#  - The Script assumes that there is a OU called Inactive in the computers OU and that the computers are moved to this OU when they are disabled.
# $activeDirectoryBaseDN - The base DN for the Active Directory domain
# $activeDirectoryComputersDN - The DN for the computers OU in Active Directory
# $computersInactiveAfter - The number of days after which a computer is considered inactive
# $technicianToEmail - The email address of the technician to send the email to
# $defaultSMTPFromAddress - The default email address to use for sending emails
# $smtpServerAddress - The SMTP server to use for sending emails
# ==============================================================================================


#Validate and Set the computers OU
$computersOU = if ($activeDirectoryComputersDN -notlike "$activeDirectoryBaseDN%") {
  "$activeDirectoryComputersDN,$activeDirectoryBaseDN"
} else {
  $activeDirectoryComputersDN
}

#Pull todays date for comparison operations
$today=(Get-Date -Format dd/MM/yyyy)

$moveResults = @()

function Lock-inactiveDevice ($inactivePC)
{
try {

      # Get current location for tracking
      $CurrentOU = ($inactivePC.DistinguishedName -split ',', 2)[1]

      # Move the computer to the new OU
      Move-ADObject -Identity $inactivePC.DistinguishedName -TargetPath "OU=Inactive,$computersOU"  -Credential $secret:computerManagement -ErrorAction SilentlyContinue
      Set-ADComputer -Identity $inactivePC.name -enabled $false  -Credential $secret:computerManagement
      Set-ADComputer -Identity $inactivePC.name -Replace @{Description=$("Computer Disabled for inactivity ($today)")}  -Credential $secret:computerManagement

      # Create a custom object to track the move
      $MoveInfo = [PSCustomObject]@{
          Name   = $inactivePC.Name
          OldOU  = $CurrentOU
          Status = "Success"
      }
      #Write-Host "Successfully moved computer '$($inactivePC.Name)'" -ForegroundColor Green
  }
  catch {
      # Create a custom object for failed moves
      $MoveInfo = [PSCustomObject]@{
          Name   = $inactivePC.Name
          OldOU  = if ($null -ne $inactivePC) { ($inactivePC.DistinguishedName -split ',', 2)[1] } else { "Unknown" }
          Status = "Failed: $_"
      }

      #Write-Host "Failed to move computer '$($inactivePC.Name)'. Error: $_" -ForegroundColor Red
  }

  # Add the move info to our results array
  return $MoveInfo
}


function Send-emailNotification($To, $From, $SMTPServer, $DisabledPC)
{
  $DisabledPC=($disabledPC -join "<br>")
  $EmailBody = @"
<!DOCTYPE html>
<html>
<head>
<style>
body {
      background-color: #f6f6f6;
      font-family: sans-serif;
      -webkit-font-smoothing: antialiased;
      font-size: 14px;
      line-height: 1.4;
      margin: 0;
      padding: 0;
      -ms-text-size-adjust: 100%;
      -webkit-text-size-adjust: 100%; }

.wrapper {
      box-sizing: border-box;
      padding: 20px;
}
.main {
      background: #fff;
      border-radius: 3px;
      width: 100%; }

    /* -------------------------------------
        TYPOGRAPHY
    ------------------------------------- */
    h1,
    h2 {
font-size: 20px;
   color: #247454;
      font-family: sans-serif;
      font-weight: 500;
      line-height: 1;
      margin: 0;
      Margin-bottom: 30px;
}
 }
    h3,
    h4 {
      color: #000000;
      font-family: sans-serif;
      font-weight: 500;
      line-height: 1.4;
      margin: 0;
      Margin-bottom: 30px; }
    h1 {
      font-size: 25px;
      font-weight: 300;
      text-align: center;
      text-transform: capitalize; }
    p,
    ul,
    ol {
      font-family: sans-serif;
      font-size: 14px;
      font-weight: normal;
      margin: 0;
      Margin-bottom: 15px; }
      p li,
      ul li,
      ol li {
        list-style-position: inside;
        margin-left: 5px; }
    a {
      color: #3498db;
      text-decoration: underline; }

</style>
</head>
<body>
<table class="main">
            <tr>
              <td class="wrapper">
                <table border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td>
                      <p><h1>Inactive Computer Accounts have been Disabled.</h1></p>

                            </td>
                          </tr>
                        </tbody>
                      </table>
                      <p>
<p><h2>Disabled Device(s):</h2> <ol>$DisabledPC</ol></p>

                    </td>
                  </tr>
                </table>
              </td>
            </tr>

</body>
</html>
"@
Send-MailMessage -to "$To" -from "$From" -subject "Disabled Computer Accounts"  -SmtpServer "$SMTPServer" -Body "$EmailBody" -bodyashtml -WarningAction Ignore
}


$disabledcomputers=@()
$inactiveTimestamp = ((Get-Date).Adddays(-($computersInactiveAfter))).ToFileTime()
foreach ($inactiveDevice in (Get-ADComputer -Filter {LastLogonTimeStamp -lt $inactiveTimestamp}  -Properties LastLogonTimeStamp -SearchBase $computersOU -Credential $secret:computerManagement | select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}},enabled,DistinguishedName | Where-Object {($_.enabled -eq $true)}))
{
$moveResults += Lock-inactiveDevice ($inactiveDevice)
$disabledcomputers += "Hostname: $($inactiveDevice.Name) <br> Last Active: $(Get-Date $inactiveDevice.Stamp -format "yyyy-MM-dd HH:MM:ss") <br>"
}



if ($disabledcomputers.Length -gt 0)
{
foreach ($emailAddress in $technicianToEmail)
{
  Send-emailNotification $emailAddress $defaultSMTPFromAddress $smtpServerAddress $disabledcomputers
}
 # Display the results in a table
Write-Output  $moveResults | Format-Table -AutoSize
}
elseif ($disabledcomputers.Length -eq 0)
{
Write-Output "No devices to disable"
}