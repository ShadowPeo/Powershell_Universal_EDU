# ==============================================================================================
# NAME: sync-eduhub-apple_school_manager.ps1
# DATE  : 01/09/2025
#
# COMMENT: Converts EduHub SF table (csv file) into standard ASM import file
# VERSION: 1, Migration to Powershell Universal from Powershell Core
# ==============================================================================================


# ==============================================================================================
# This script is designed to be run on Powershell Universal not directly as a script and is looking for the following variables to be set in the environment set by the server.
#  - As this uses active directory it is assumed that the server is a member of the domain and has access to the Active Directory servers. and the Active Directory module is installed.
#  - The custom module JAMF.ps1 is also required to be in the same directory as this script this will be migrated at some point to either a full module installed, an existing JAMF Module that can be installed or a set of API endpoints that can be used to get the data from JAMF from the PSU server.
# $activeDirectorySchoolServers - Array of Active Directory servers to use for the school, minimum of 1 required. - This will be migrated to a lookup at some point
# $schoolDomainSuffix - The domain suffix to use for the school. - This will be migrated to a lookup at some point
# $eduhubDataPath - The path to the EduHub data files.
# $jamfTennant - The Jamf Pro URL to use for the school.
# $jamfCredentials - The Jamf Pro credentials to use for the school. - Set as a secret in the environment using PSCredential format
# $appleSMServer - The Apple School Manager SFTP server to use for the school.
# $appleSFTPDetails - The Apple School Manager credentials to use for the school. - Set as a secret in the environment using PSCredential format.
# $appleLocationID - The Apple School Manager location ID to use for the school.

# ==============================================================================================

### Run
$clearFiles = $false #deletes temp files from working folders after running.
$dryRun = $false # A dry run will not attempt to upload to the Apple FTP
$Primary = $true #Primary school?
$Secondary = $false #Secondary school? - the difference is the classes are not simplified.
#$promotionCurrent = $true
$outputYear = "2025"

#Local AD Settings
$adServer = $activeDirectorySchoolServers[0]
#Folder\working paths
$ASMfolderpath = "C:\Scripts\DataSync\ASM\"
#$LogFile = "$ASMfolderpath\Logs\$(Get-Date -UFormat '+%Y-%m-%d-%H-%M')-$(if($dryRun -eq $true){"DRYRUN-"})$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log"

### Apple School Manager settings
$STUPWPOL = "" # Student Password Policy for ASM "4,6 or 8"

#Import Modules
Import-Module "$ASMfolderpath/Modules/JAMF.ps1"
Import-Module ActiveDirectory

$token = Get-jamfToken -tennantURL $jamfTennant -user $($Secret:jamfCredentials.username) -pass $((New-Object PSCredential 0, $Secret:jamfCredentials.password).GetNetworkCredential().Password)
$iPadsBYOD = (Get-jamfMobileDevices -tennantURL $jamfTennant -token $token).results | Where-Object -FilterScript {($_.model -like "iPad*") -and -not [string]::IsNullOrWhiteSpace($_.username) -and -not [string]::IsNullOrWhiteSpace($_.name)} | Sort-Object username | Where-Object name -like "iPadB*" | Select-Object username


# do not change below this line unless you know what you are doing.

 # this alters the domain suffix for staff and students email accounts (only if integrated)
 # It's not currently used actively as it builds the email address incorrectly for both '@schools' student accouts and '@localschooldomain.vic.edu.au' staff... for now.

 $EDU001_integrated = $false
# if integrated is set to true, the script will set the student *email* account to @schools.vic.edu.au but leave the student *apple ID* suffix as @localschooldomain.vic.edu.au
# if integrated is set to true, the script will set the staff *email* account to @education.vic.gov.au but leave the staff *apple ID* suffix as @localschooldomain.vic.edu.au
# if integrated is set to false, the script will set the student *email* account to @localschooldomain.vic.edu.au AND leave the student *apple ID* suffix as @localschooldomain.vic.edu.au
# if integrated is set to false, the script will set the staff *email* account to @localschooldomain.vic.edu.au AND leave the staff *apple ID* suffix as @localschooldomain.vic.edu.au

if ($EDU001_integrated) {

    $Staff_emailID_Domain_Suffix = "education.vic.gov.au"

    $Student_emailID_Domain_Suffix = "schools.vic.edu.au"

}else{

    $Staff_emailID_Domain_Suffix = "$schoolDomainSuffix"

    $Student_emailID_Domain_Suffix = "$schoolDomainSuffix"

}

# test for 'working' and 'final' folders in the scripts '$ASMfolderpath' directory

If(!(test-path $ASMfolderpath))
{
    New-Item -ItemType Directory -Force -Path $ASMfolderpath
}

$ASMworkingpath = $ASMfolderpath + "Working\"
if(!(test-path $ASMworkingpath))
{
    New-Item -ItemType Directory -Force -Path $ASMworkingpath
}

$ASMfinalpath = $ASMfolderpath + "Final\"
if(!(test-path $ASMfinalpath))
{
    New-Item -ItemType Directory -Force -Path $ASMfinalpath
}

if(!(test-path "$ASMfolderpath\Logs"))
{
    New-Item -ItemType Directory -Force -Path "$ASMfolderpath\Logs"
}

#Write-Output $PSVersionTable.PSVersion

$TimeStamp = Get-Date -Format MM-dd-yyyy_HH_mm_ss
$finalArchive = "$ASMfinalpath" + "asmUpdate-" + "$TimeStamp.zip"

# set eduhub and asm file location variables

$OutPutStaffFile = $ASMworkingpath + "staff.csv"
$OutPutStudentFile = $ASMworkingpath + "students.csv"
$OutPutClassesFile = $ASMworkingpath + "classes.csv"
$OutPutCoursesFile = $ASMworkingpath + "courses.csv"
$OutPutLocationsFile = $ASMworkingpath + "locations.csv"
$OutPutRostersFile = $ASMworkingpath + "rosters.csv"

### Env

$SchoolStaffFile = "$eduhubDataPath\" + "CasesStaff-NEWFORM.csv"
$SchoolStudentFile = "$eduhubDataPath\" + "CasesStudents-NEWFORM.csv"
$SchoolClassesFile = "$eduhubDataPath\" + "SCL_$schoolNumber.csv"
$SchoolCoursesFile = "$eduhubDataPath\" + "SU_$schoolNumber.csv"
$SchoolLocationsFile = "$eduhubDataPath\" + "SCI_$schoolNumber.csv"
$SchoolRostersFile = "$eduhubDataPath\" + "STMA_$schoolNumber.csv"
$SchoolKGCFile = "$eduhubDataPath\" + "KGC_$schoolNumber.csv"



# in order to export the whole school in a multicampus school remove the line "	where-object{$_.CAMPUS -eq "1"}| and CAMPUS = Apples "location_id"
# This not required in a single campus school or when exporting the campuses separately.
Write-Output "Reading STAFF eduhub CSV and pull ASM required details to build ASM valid CSV..."
## Read STAFF eduhub CSV and pull ASM required details to build ASM valid CSV
$tempStaff = Import-Csv $SchoolStaffFile |
	where-object{($_.SIS_EMPNO -ne "") -and ($_.STATUS -eq "ACTV")} |
    Select-Object @{Name="person_id";Expression={$_."SIS_EMPNO"}}, @{Name="person_number";Expression={$_."SIS_ID"}}, @{Name="first_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."FIRST_NAME".ToLower())}}, @{Name="middle_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."SECOND_NAME".ToLower())}}, @{Name="last_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."SURNAME".ToLower())}}, @{Name="email_address";Expression={ if ($_.E_MAIL) { $_.E_MAIL } else { (Get-Culture).TextInfo.ToTitleCase($_."FIRST_NAME".ToLower()) + '.' + (Get-Culture).TextInfo.ToTitleCase($_."SURNAME".ToLower()) + '@' + $Staff_emailID_Domain_Suffix } } }, @{Name="sis_username";Expression={$_."SIS_EMPNO"}}, @{Name="location_id";Expression={"$appleLocationID"}} |

    Sort-Object -property person_id

$outputStaff = @()

foreach ($staffMember in $tempStaff)
{
    $adUser = $null
    if ($staffMember.person_number -ne "" -and $null -ne $staffMember.person_number)
    {
        try {
            # Try to get AD user
            $adUser = Get-ADUser -identity $staffMember.person_id -Server $adServer -Properties mail
            Write-Output "User found: $($adUser.SamAccountName)"
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Output "User not found in Active Directory"
        }
        catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
            Write-Output "Cannot connect to Active Directory server"
            Write-Output "Error Message: $($_.Exception.Message)"
            Write-Output "Error details: $($_.Exception.GetType().FullName)"
        }
        catch [System.UnauthorizedAccessException] {
            Write-Output "Access denied - insufficient permissions"
        }
        catch {
            Write-Output "Unexpected error: $($_.Exception.Message)"
            Write-Output "Error details: $($_.Exception.GetType().FullName)"
        }

    }
    else
    {
        $adUser = Get-ADUser -Server $adServer -Identity $staffMember.person_number -Properties mail -ErrorAction SilentlyContinue
    }

    if (($null -ne $adUser.mail -and  $adUser.mail -ne "") -and $null -ne $adUser)
    {
        $staffMember.email_address = ($adUser.mail).ToLower()
        $outputStaff += $staffMember
    }
}
# convert to csv and remove pipes and double quotes.
    $outputStaff |
    ConvertTo-Csv -NoTypeInformation |
    ForEach-Object {$_.Replace('|','-')} |
    Out-File $OutPutStaffFile -encoding ascii


## Read STUDENT eduhub CSV and pull ASM required details to build ASM valid CSV
Import-Csv $SchoolStudentFile |
    where-object{($_.SIS_ID -ne "") -and ($_.STATUS -eq "FUT" -or "ACTV" -or "LVNG") -and ($_.STATUS -ne "LEFT") -and ($_.STATUS -ne "DEL") -and ($_.CAMPUS -ne "")} |
    Select-Object @{Name="person_id";Expression={($_."SIS_ID").ToLower()}}, @{Name="person_number";Expression={$_."SIS_ID"}}, @{Name="first_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."FIRST_NAME".ToLower())}}, @{Name="middle_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."SECOND_NAME".ToLower())}}, @{Name="last_name";Expression={(Get-Culture).TextInfo.ToTitleCase($_."SURNAME".ToLower())}}, @{Name="grade_level";Expression={$_."SCHOOL_YEAR"}}, @{Name="email_address";Expression={$_.SIS_ID.ToLower() + '@' + $Student_emailID_Domain_Suffix}}, @{Name="sis_username";Expression={$_."SIS_ID"}}, @{Name="password_policy";Expression={$STUPWPOL}}, @{Name="location_id";Expression={"$appleLocationID"}} |
	Sort-Object -property person_id |
# convert to csv and remove pipes and double quotes.
 	ConvertTo-Csv -NoTypeInformation |
    ForEach-Object {$_.Replace('|','-')} |
    Out-File $OutPutStudentFile -encoding ascii

## LOCATION.
# Read school eduhub CSV and pull ASM required details to build ASM valid CSV
Import-Csv $SchoolLocationsFile |

	Select-Object @{Name="location_id";Expression={"$appleLocationID"}}, @{Name="location_name";Expression={$_."CAMPUS_NAME"+' - Campus '+$_."SCIKEY" }} |
    Sort-Object -property location_id |
# convert to csv and remove pipes and double quotes.
 	ConvertTo-Csv -NoTypeInformation |
    ForEach-Object {$_.Replace('|','-')} |
    Out-File $OutPutLocationsFile -encoding ascii


if ($Primary)
{
    ##CLASSES (KGC)- The specific class: Mr Smith Algebra III 5th period
    # Read school eduhub CSV and pull ASM required details to build ASM valid CSV
    $tempClasses = Import-Csv $SchoolKGCFile |
    where-object{($_.ACTIVE -eq "Y")} |
        Select-Object @{Name="class_id";Expression={"$($_."KGCKEY")-$outputYear"}}, @{Name="class_number";Expression={$_."KGCKEY"}}, @{Name="course_id";Expression={$_."KGCKEY"}}, @{Name="instructor_id";Expression={(Get-Culture).TextInfo.ToTitleCase($_."TEACHER")}}, @{Name="instructor_id_2";Expression={(Get-Culture).TextInfo.ToTitleCase($_."TEACHER_B")}}, @{Name="instructor_id_3";Expression={''}}, @{Name="location_id";Expression={$appleLocationID}} |
        Sort-Object -property class_id

    foreach ($class in $tempClasses)
    {
    if ((($class.course_id).Substring(0,2)) -eq "0F")
    {
            $class.course_id = "00"
    }
    elseif ((($class.course_id).Substring(0,2)) -eq "ZZ")
    {
            $class.course_id = "ZZZ"
    }
    elseif ((($class.course_id).Substring(0,2)) -ge "01" -and (($class.course_id).Substring(0,2)) -lt "12")
    {
            $class.course_id = (($class.course_id).Substring(0,2))
    }

    if($class.instructor_id -ne "" -and $null -ne $class.instructor_id)
    {
            $tempStaff = $null
            $tempStaff = $outputStaff | Where-Object {($_.person_number -eq $class.instructor_id)}
            $class.instructor_id = $tempStaff.person_id
    }
    if($class.instructor_id_2 -ne "" -and $null -ne $class.instructor_id_2)
    {
            $tempStaff = $null
            $tempStaff = $outputStaff | Where-Object {($_.person_number -eq $class.instructor_id_2)}
            $class.instructor_id_2 = $tempStaff.person_id
    }
    if($class.instructor_id_3 -ne "" -and $null -ne $class.instructor_id_3)
    {
            $tempStaff = $null
            $tempStaff = $outputStaff | Where-Object {($_.person_number -eq $class.instructor_id_3)}
            $class.instructor_id_3 = $tempStaff.person_id
    }

    }
    $tempClasses | Sort-Object course_id, class_number |
    # convert to csv and remove pipes and double quotes.
        ConvertTo-Csv -NoTypeInformation |
        ForEach-Object {$_.Replace('|','-')} |
        Out-File $OutPutClassesFile -encoding ascii


    ##COURSES (KGC)- Course are the generic group: Algebra III - North High School
    # Read school eduhub CSV and pull ASM required details to build ASM valid CSV
    $tempCourses = Import-Csv $SchoolStudentFile |
                    where-object{($_.SIS_ID -ne "") -and ($_.STATUS -ne "LEFT") -and ($_.STATUS -ne "DEL")} |
                    Sort-Object -Unique -Property "SCHOOL_YEAR" |
                    Select-Object @{Name="course_id";Expression={$_."SCHOOL_YEAR"}},@{Name="course_number";Expression={$_."SCHOOL_YEAR"}}, @{Name="course_name";Expression={"Year $($_."SCHOOL_YEAR")"}}, @{Name="location_id";Expression={"$appleLocationID"}}
    # Add ZZZ Course for prestaging future students
    $tempOutput = New-Object PSObject
    $tempOutput | Add-Member -MemberType NoteProperty -Name "course_id" -Value ("ZZZ")
    $tempOutput | Add-Member -MemberType NoteProperty -Name "course_number" -Value ("ZZZ")
    $tempOutput | Add-Member -MemberType NoteProperty -Name "course_name" -Value ("Future Students")
    $tempOutput | Add-Member -MemberType NoteProperty -Name "location_id" -Value ($appleLocationID)
    $tempCourses += $tempOutput

    # convert to csv and remove pipes and double quotes.
    $tempCourses | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_.Replace('|','-')} | Out-File $OutPutCoursesFile -encoding ascii

    ##ROSTERS (ST)- Each student that attends each class: Roster ID - Class ID - Student ID
    ##Since the Roster ID is unique, a student would have a row for each class they are taking. A student in high school may have 7-8 rows. A class may have 30 entries, one for each student.
    Import-Csv $SchoolStudentFile | where-object -FilterScript {($_.SIS_ID -ne "") -and ($_.STATUS -ne "LEFT") -and ($_.STATUS -ne "DEL") -and ($iPadsBYOD.Username -contains "$($_.SIS_ID)@$schoolDomainSuffix") } | Select-Object @{Name="roster_id";Expression={"$($_."SIS_ID")-$($_."HOME_GROUP")-$outputYear"}}, @{Name="class_id";Expression={"$($_."HOME_GROUP")-$outputYear"}}, @{Name="student_id";Expression={($_."SIS_ID").ToLower()}} | Sort-Object -property roster_id | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_.Replace('|','-')} | Out-File $OutPutRostersFile -encoding ascii
}

if ($Secondary)
{
    ##CLASSES (SCL)- The specific class: Mr Smith Algebra III 5th period
    Import-Csv $SchoolClassesFile |
        Select-Object @{Name="class_id";Expression={$_."SCLKEY"}}, @{Name="class_number";Expression={$_."CLASS"}}, @{Name="course_id";Expression={$_."SUBJECT"}}, @{Name="instructor_id";Expression={(Get-Culture).TextInfo.ToTitleCase($_."TEACHER01")}}, @{Name="instructor_id_2";Expression={(Get-Culture).TextInfo.ToTitleCase($_."TEACHER02")}}, @{Name="instructor_id_3";Expression={''}}, @{Name="location_id";Expression={$_."CAMPUS"}} |
        Sort-Object -property class_id  |
    # convert to csv and remove pipes and double quotes.
        ConvertTo-Csv -NoTypeInformation |
        ForEach-Object {$_.Replace('|','-')} |
        Out-File $OutPutClassesFile -encoding ascii

    ##COURSES (SU)- Course are the generic group: Algebra III - North High School
    Import-Csv $SchoolCoursesFile |
        Select-Object @{Name="course_id";Expression={$_."SUKEY"}}, @{Name="course_number";Expression={$_."SUKEY"}}, @{Name="course_name";Expression={$_."FULLNAME"}}, @{Name="location_id";Expression={''}} |
        Sort-Object -property course_id  |
    # convert to csv and remove pipes and double quotes.
        ConvertTo-Csv -NoTypeInformation |
        ForEach-Object {$_.Replace('|','-')} |
        Out-File $OutPutCoursesFile -encoding ascii

    ##ROSTERS (STMA)- Each student that attends each class: Roster ID - Class ID - Student ID
    ##Since the Roster ID is unique, a student would have a row for each class they are taking. A student in high school may have 7-8 rows. A class may have 30 entries, one for each student.
    Import-Csv $SchoolRostersFile |
        Select-Object @{Name="roster_id";Expression={$_."TID"}}, @{Name="class_id";Expression={$_."CKEY"}}, @{Name="student_id";Expression={$_."SIS_ID"}} |
        Sort-Object -property roster_id |
    # convert to csv and remove pipes and double quotes.
        ConvertTo-Csv -NoTypeInformation |
        ForEach-Object {$_.Replace('|','-')} |
        Out-File $OutPutRostersFile -encoding ascii
}

Write-Output "ASM CSVs created."

####### Generate ZIP archive #######
Write-Output "Zipping up ASM CSVs..."
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($ASMworkingpath, $finalArchive)
Write-Output "Done"
#Stop-Transcript
####### Manage dry run if needed #######

if ($dryRun) {
    Write-Output "CSVs and zip files created. It's a dry run so no upload to Apple School Manager."
    Write-Output "$ASMworkingpath"
    Invoke-Item "$ASMworkingpath"
    exit 0
}

###SFTP Session

#Main Code
#Change the location to the same directory of WinSCP assembly avoids register the dll file on the server

if (!$dryRun)
{
    #Check Putty SCP exists, if not attempt to download it
    if (!(Test-Path "$ASMfolderpath/pscp.exe" -PathType Leaf))
    {
        Write-Output "PSCP not found, downloading"
        try
        {
            if (-not ([string]::IsNullOrWhiteSpace($proxyAddress)))
            {
                Invoke-WebRequest -Uri "https://the.earth.li/~sgtatham/putty/latest/w64/pscp.exe" -OutFile "$ASMfolderpath/pscp.exe" -Proxy $proxyAddress  | Out-Null
            }
            else
            {
                Invoke-WebRequest -Uri "https://the.earth.li/~sgtatham/putty/latest/w64/pscp.exe" -OutFile "$ASMfolderpath/pscp.exe"  | Out-Null
            }
        }
        catch
        {
            $_.Exception.Response.StatusCode.Value__
        }
    }
    else
    {
        Write-Output "PSCP Found, Continuing"
    }
    & "$ASMfolderpath\pscp.exe" -q -pw $((New-Object PSCredential 0, $Secret:appleSFTPDetails.password).GetNetworkCredential().Password) -sftp -l $($Secret:appleSFTPDetails.username) $finalArchive ("$appleSMServer`:/dropbox/" + "asmUpdate-$TimeStamp.zip")
}

if ($clearFiles) {
    Write-Output "removing $finalArchive"
    Remove-Item -path "$finalArchive"
    exit 0
}

exit 0