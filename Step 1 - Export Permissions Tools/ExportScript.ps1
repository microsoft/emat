Param(
  [bool]$silent=$false
)


# Prerequisites Check


if ((get-host).Version.Major -lt 4) {

    write-host -foregroundcolor red "You are running an old (and unsupported) version of Powershell, please consider to update the Powershell before running this script"
    write-host "Press any key to continue..."
    [void][System.Console]::ReadKey($true)	
    exit

}

if ( (Get-Command get-mailbox -errorAction SilentlyContinue) -eq $null )
{
    write-host -foregroundcolor red "Exchange Powershell module is not available, be sure to run the script from a server that have this feature installed"
    write-host "Press any key to continue..."
    [void][System.Console]::ReadKey($true)
    exit
}


if ( (Get-Command get-group -errorAction SilentlyContinue) -eq $null )
{
    write-host -foregroundcolor red "AD DS Powershell module is not available, be sure to run the script from a server that have this feature installed"
    write-host "Press any key to continue..."
    [void][System.Console]::ReadKey($true)
    exit
}



# To run the script as a scheduled task or unattend script uncomment the following line
# $silent=$true

# Set the PowerShell to search Forest Wide Mailboxes
Set-ADServerSettings -ViewEntireForest $True

# Get the list of mailboxes with required attribute
get-mailbox * -resultsize unlimited | select Identity,Name,DisplayName,samaccountname,RecipientType,RecipientTypeDetails,WindowsEmailAddress | export-csv mailboxes.csv -encoding "unicode" -NoTypeInformation

# Create multiple sessions to speed up the export
if ($silent -eq $false) {
	$MailboxesFileSize=[math]::truncate((Get-Item .\mailboxes.csv).Length /1024 /1024)
	write-host -ForegroundColor green "Current mailboxes.csv size is $MailboxesFileSize MB"
	}
if ($MailboxesFileSize -gt 7 -and $silent -eq $false) {
	write-host -ForegroundColor green "I think that I can export permissions of about 7 MB users file in 24 hours (this is just a statistic point of view)"
	$response=read-host "Do you want to split the mailboxes file? (Y/N)"

	if ($response -eq "y" -or $response -eq "Y" -or $response -eq "") {
		$splitrow=Get-Content .\ExportScript.ps1 | Select-String "# GOTO::START OF PERMISSIONS EXPORT" | Select-Object -ExpandProperty 'LineNumber'
		write-host -ForegroundColor green "I'll do the split for you"
		$howmany=read-host "How many sessions/servers you want to use?"
		write-host -ForegroundColor green -NoNewLine "Reading mailboxes file..."
		$original_file=import-csv .\mailboxes.csv
		$numberofrows=$original_file.Count
		write-host -ForegroundColor yellow "DONE"

		$block=[math]::truncate($numberofrows / $howmany)
		$first=1
		$last=$block
		mkdir Sessions > $null
		cd Sessions
		mkdir Session1  > $null
		$original_file | select -first $first -last $last | export-csv Session1\mailboxes.csv -encoding "unicode" -NoTypeInformation
		Get-Content ..\ExportScript.ps1 | select -skip $splitrow[1] >> Session1\ExportScript.ps1
		write-host -foregroundColor green "The files will cointains about $block users each"
		write-host -ForegroundColor yellow "Session 1 - Created - from $first to $last"

		for ($i=2;$i -le $howmany;$i++) {
			$first=$first+$block
			$last=$last+$block
			mkdir Session$i > $null
			$original_file | select -skip $first -first $block | export-csv Session$i\mailboxes.csv -encoding "unicode" -NoTypeInformation
			Get-Content ..\ExportScript.ps1 | select -skip $splitrow[1] >> Session$i\ExportScript.ps1
			write-host -ForegroundColor yellow "Session $i - Created - from $first to $last"
		}
	}
	cd ..
	write-host -foregroundColor green "Sessions created, I'll go ahead with groups export while you can copy the sessions folders on the servers that will run it"
}



# Get the list of the groups, only the MailUniversalSecurityGroup (Exchange supported groups to provide permissions) are fetched
Get-Group -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Select Name,Alias,DisplayName,WindowsEmailAddress,SamAccountName | export-csv groups.csv -NoTypeInformation

# Using the list of groups export the group membership
$groups=(import-csv groups.csv).samaccountname 
$table=New-Object system.Data.DataTable "OutputGroups"
$col1=New-Object system.data.dataColumn SamAccountNAme,([string])
$col2=New-Object system.data.dataColumn DGName,([string])
$col3=New-Object system.data.dataColumn DomainNB,([string])
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)

foreach ($group in $groups) {
	$asd=Get-ADGroupMember $group 
	foreach ($user in $asd) {
			$row = $table.NewRow()
			$row.SamAccountNAme=$user.SamAccountname
			$row.DGName=$group 
			$row.DomainNB="DUMMY"
			$table.rows.add($row)
		}
	}
$table | export-csv groupsmembership.csv -encoding "unicode" -NoTypeInformation 
# End of Group Membership Block



if ($MailboxesFileSize -gt 7 -and $silent -eq $false) {
	write-host -foregroundColor yellow "I created splitted session previously so please go ahead running single sessions"
	exit
}


# GOTO::START OF PERMISSIONS EXPORT
# Using the mailboxes file, search for permissions for each mailbox
# NOTE: at this point you can split the mailboxes file, copying the first row with headers in each file, and execute the scripts copy in paralell
# on different servers/PowerShell sessions to speed up the execution


# ADPermissions: Send-AS, Receive-AS
import-csv mailboxes.csv | Get-ADPermission | select Identity,User,@{Name='AccessRight';Expression={"AD"}},ExtendedRight,FolderPath | export-csv ad.csv -encoding "unicode" -NoTypeInformation 

# Mailbox Permissions: FullAccess, Read
import-csv mailboxes.csv | Get-MailboxPermission | select Identity,User,@{Name='AccessRight';Expression={"MP"}},ExtendedRight,FolderPath | export-csv mp.csv -encoding "unicode" -NoTypeInformation 

# SendOnBeHalf permissions
# import-csv mailboxes.csv | get-mailbox | ?{$_.GrantSendOnBehalfTo} | select Identity,@{Name='User';Expression={[string]::join(";",($_.GrantSendOnBehalfTo))}},@{Name='AccessRight';Expression={"SNB"}},ExtendedRight,FolderPath | export-csv SnB.csv -encoding "unicode" -NoTypeInformation 

# SendOnBeHalf permissions (NEW VERSION)
import-csv mailboxes.csv | get-mailbox | ?{$_.GrantSendOnBehalfTo} | select Identity,@{Name='User';Expression={$_.GrantSendOnBehalfTo -join ";"}},@{Name='AccessRight';Expression={"SNB"}},ExtendedRight,FolderPath | export-csv SnB.csv -encoding "unicode" -NoTypeInformation


# Get Folder Permissions (Exchange 2010+ is required)
# This section of the script may generate a lot of errors due no informations for system folders like "\Version"

$table=New-Object system.Data.DataTable "Output"
$col1=New-Object system.data.dataColumn Identity,([string])
$col2=New-Object system.data.dataColumn User,([string])
$col3=New-Object system.data.dataColumn PermissionType,([string])
$col4=New-Object system.data.dataColumn AccessRights,([string])
$col5=New-Object system.data.dataColumn ExtendedRight,([string])
$col6=New-Object system.data.dataColumn FolderPath,([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)

$all=import-csv mailboxes.csv
foreach ($identity in $all) {
        $i=$i+1
        [System.Collections.ArrayList]$folders = Get-MailboxFolderStatistics $identity.Identity | % {$_.folderpath} | % {$_.replace("/","\")}
	$alias=$identity.windowsEmailAddress
	$folders.Remove("\Versions")
	$folders.Remove("\Journal")	
	$folders.Remove("\Deletions")
	$folders.Remove("\Calendar Logging")
	$folders.Remove("\Recoverable Items")
	$folders.Remove("\Livello superiore archivio informazioni")
	$folders.Remove("\Working Set")
	$folders.Remove("\Posta indesiderata")
	$folders.Remove("\Conversation Action Settings")
	$folders.Remove("\Purges")
	$folders.Remove("\Posta in uscita")
	$folders.Remove("\Contatti\Recipient Cache")
	$folders.Remove("\Top of Information Store")


        $utente = $folders | %{ Get-MailboxFolderPermission $alias":"$_ | select Identity,User,Accessrights,foldername} | where {$_.User -notlike "Default*" -and $_.User -notlike "Anonymous*"}
	

	foreach ($riga in $utente) {
		$acc=[string]::join(";",($riga.Accessrights))
		$row = $table.NewRow()
		$row.Identity=$identity.Identity
		$row.User=$riga.User
		$row.AccessRights="FP"
		$row.PermissionType="FP"
		$row.ExtendedRight=""
		$row.FolderPath=$riga.foldername
		$table.rows.add($row)
		}
}

$table | export-csv fp.csv -encoding "unicode" -NoTypeInformation 
# End of get Folder Permissions block