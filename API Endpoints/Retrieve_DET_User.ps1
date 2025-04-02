$userIdentity = $null
$userEmail = $null
$userName = $null
$userType = $null


Import-Module ActiveDirectory
$userToReturn = $null

#Lookup Based upon idenity provided as first preference
if (![string]::IsNullOrWhiteSpace($userIdentity))
{
        try
        {
            $userToReturn =  Get-ADUser -Server $edu002ADServer -Credential $Secret:edu002DCUser -Identity $userIdentity -Properties $studentDETDetails
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {
            #Write-Warning "User '$userIdentity' not found in EDU002 Active Directory, trying EDU001"
            try
            {
                $userToReturn = Get-ADUser -Server $edu001ADServer -Credential $Secret:edu002DCUser -Identity $userIdentity -Properties $staffDETDetails
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                #Write-Warning "User '$userIdentity' not found in EDU001 Active Directory, User Cannot be found anywhere"
            }
            catch
            {
                Write-Error "An unexpected error occurred: $_"
            }
        }
        catch
        {
            Write-Error "An unexpected error occurred: $_"
        }
}

if (![string]::IsNullOrWhiteSpace($userEmail) -and $null -eq $userToReturn)
{

        if ($userEmail -ilike "*@education.vic.gov.au")
        {
            try
            {
                $userToReturn = Get-ADUser -Server $edu001ADServer -Credential $Secret:edu002DCUser -Filter {mail -eq $userEmail} -Properties $staffDETDetails
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                #Write-Warning "User Email '$userEmail' not found in EDU001 Active Directory"
            }
            catch
            {
                Write-Error "An unexpected error occurred: $_"
            }
        }
        elseif ($userEmail -ilike "*@schools.vic.edu.au")
        {
            try
            {
                $userToReturn = Get-ADUser -Server $edu002ADServer -Credential $Secret:edu002DCUser -Filter {mail -eq $userEmail} -Properties $studentDETDetails
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                #Write-Warning "User Email '$userEmail' not found in EDU001 Active Directory"
            }
            catch
            {
                Write-Error "An unexpected error occurred: $_"
            }
        }
}

if ((![string]::IsNullOrWhiteSpace($userName) -and ![string]::IsNullOrWhiteSpace($userType) )-and $null -eq $userToReturn)
{

        if ($userType -ieq "STAFF")
        {
            try
            {
                $userToReturn = Get-ADUser -Server $edu001ADServer -Credential $Secret:edu002DCUser -Filter {DisplayName -eq $userName} -SearchBase "OU=Users,OU=Schools,DC=education,DC=vic,DC=gov,DC=au" -Properties $staffDETDetails | Where-Object {$_.memberof -contains "CN=$($schoolNumber)-gs-All Staff,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                #Write-Warning "User Email '$userEmail' not found in EDU001 Active Directory"
            }
            catch
            {
                Write-Error "An unexpected error occurred: $_"
            }
        }
        elseif ($userType -ieq "STUDENT")
        {
            try
            {
                $userToReturn = Get-ADUser -Server $edu002ADServer -Credential $Secret:edu002DCUser -Filter {DisplayName -eq $userName} -SearchBase "OU=Accounts,DC=services,DC=education,DC=vic,DC=gov,DC=au " -Properties $studentDETDetails | Where-Object {$_.memberof -contains "CN=$($schoolNumber)-gs-All Students,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                #Write-Warning "User Email '$userEmail' not found in EDU001 Active Directory"
            }
            catch
            {
                Write-Error "An unexpected error occurred: $_"
            }
        }
}

if ($null -ne $userToReturn)
{
    return $userToReturn | ConvertTo-Json  -WarningAction Ignore
}