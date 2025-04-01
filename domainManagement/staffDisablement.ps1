# ==============================================================================================
# NAME: User Disablement.ps1
# DATE  : 28/03/2025
#
# COMMENT: Checks for inactive staff in Active Directory and disables them
# VERSION: 1, Migration to Powershell Universal from Powershell Core
# ==============================================================================================


# ==============================================================================================
# This script is designed to be run on Powershell Universal not directly as a script and is looking for the following variables to be set in the environment set by the server.
#  - As this uses active directory it is assumed that the server is a member of the domain and has access to the Active Directory servers. and the Active Directory module is installed.
#  - The Script assumes that there is a OU called Inactive in the staff OU and that the staff are moved to this OU when they are disabled.
# $activeDirectoryBaseDN - The base DN for the Active Directory domain
# $activeDirectoryStaffDN - The DN for the staff OU in Active Directory
# $staffInactiveAfter - The number of days after which a staff member is considered inactive
# $technicianToEmail - The email address of the technician to send the email to
# $defaultSMTPFromAddress - The default email address to use for sending emails
# $smtpServerAddress - The SMTP server to use for sending emails
# $ignoreList - A list of samAccountNames, employeeIDs, or employeeNumbers to ignore, this is an array
# ==============================================================================================

# ==============================================================================================
# TODO:
# - Check to see if user accounts have expired and if so disable them
# ==============================================================================================

#Validate and Set the Staff OU
$staffOU = if ($activeDirectoryStaffDN -notlike "$activeDirectoryBaseDN%") {
    "$activeDirectoryStaffDN,$activeDirectoryBaseDN"
} else {
    $activeDirectoryStaffDN
}

function Send-emailNotification($emailTo, $emailFrom, $SMTPServer, $disabledUsers)
{
    $disabledUsers=($disabledUsers -join "<br>")
    $emailBody = @"
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
                        <p><h1>Inactive User Accounts have been Disabled.</h1></p>

                              </td>
                            </tr>
                          </tbody>
                        </table>
                        <p>
<p><h2>Disabled Device(s):</h2> <ol>$disabledUsers</ol></p>

                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

</body>
</html>
"@
    foreach ($emailAddress in $emailTo)
    {
        Send-MailMessage -to $emailAddress -from $emailFrom -subject "Disabled User Accounts"  -SmtpServer $SMTPServer -Body $emailBody -bodyashtml -WarningAction Ignore
    }
 }


$adServers = (Get-ADDomaincontroller -Filter *)
$userStatus = @()
$disabledUsers=@()

foreach ($domainController in $adServers)
{
    $adUsers = Get-ADuser -filter * -properties employeeID,employeeNumber,lastLogon,whenCreated,enabled -Server $domainController.hostname -SearchBase $staffOU | Select-Object samAccountName,DistinguishedName,employeeID,employeeNumber,lastLogon,whenCreated,enabled
    foreach ($adUser in $adUsers)
    {
        $retrievedUserInfo = [PSCustomObject]@{
            domainController = $domainController.name
            samAccountName= $adUser.samAccountName
            distinguishedName = $adUser.DistinguishedName
            employeeID= $adUser.employeeID
            employeeNumber= $adUser.employeeNumber
            lastLogon = $adUser.lastLogon
            whenCreated = $adUser.whenCreated
            enabled= $adUser.enabled
        }
        $userStatus += $retrievedUserInfo
    }
}

$userStatus = $userStatus | Where-Object { ($_.enabled -eq $true) } | Sort-Object samAccountName, lastLogon | Sort-Object samAccountName -Unique

foreach ($user in $userStatus)
{
    #IF user is not in the ignorelist (check samAccountName, employeeID, employeeNumber)
    if ($ignoreList -notcontains $user.samAccountName -and $ignoreList -notcontains $user.employeeID -and $ignoreList -notcontains $user.employeeNumber)
    {
        if ($null -eq $user.lastLogon -or $user.lastLogon -eq 0)
        {
            #Write-Output "User $($user.samAccountName) has never logged on"
            #Write-Output "User $($user.samAccountName) was created at $($user.whenCreated)"
            if ($user.whenCreated -lt (Get-Date).AddDays(-($staffInactiveAfter)))
            {
                Write-Output "User $($user.samAccountName) was created more than $staffInactiveAfter days ago, Disabling account"
                $disabledUsers += "Username: $($user.samAccountName) <br> Never Activated <br> Created On: $(Get-Date $user.whenCreated -format "yyyy-MM-dd HH:MM:ss") <br>"
                Set-ADUser -Identity $user.samAccountName -Description "Inactive account, disabled by script $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"
                Move-ADObject -Identity $user.DistinguishedName -TargetPath "OU=Inactive,$staffOU" -ErrorAction SilentlyContinue
                Set-ADUser -Identity $user.samAccountName -Enabled $false
            }
            else
            {
                #Write-Output "User $($user.samAccountName) was created less than $staffInactiveAfter days ago, Ignoring"
            }
        }
        else
        {
            $correctedLastLogin = $null
            $correctedLastLogin= [DateTime]::FromFileTime($user.lastLogon)
            #Write-Output "User $($user.samAccountName) logged on most recently at $(Get-Date $correctedLastLogin -Format "yyyy-MM-dd HH:MM:ss" ) according to $($user.domainController)"
            if ($correctedLastLogin -lt (Get-Date).AddDays(-$staffInactiveAfter))
            {
                Write-Output "User $($user.samAccountName) last logged on more than $staffInactiveAfter days ago, Disabling account"
                $disabledUsers += "Username: $($user.samAccountName) <br> Last Log On: $(Get-Date $correctedLastLogin -format "yyyy-MM-dd HH:MM:ss") <br>"
                Set-ADUser -Identity $user.samAccountName -Description "Inactive account, disabled by script $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"
                Move-ADObject -Identity $user.DistinguishedName -TargetPath "OU=Inactive,$staffOU" -ErrorAction SilentlyContinue
                Set-ADUser -Identity $user.samAccountName -Enabled $false
            }
            else
            {
                #Write-Output "User $($user.samAccountName) last logged on less than $staffInactiveAfter days ago, Ignoring"
            }
        }
    }
}

if ($disabledUsers.Length -gt 0)
{
    Send-emailNotification $technicianToEmail $defaultSMTPFromAddress $smtpServerAddress $disabledUsers
}