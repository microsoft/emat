# Proof of concept of a porting of the script to allow to get Exchange permissions from Office 365.
# This scenario cames up from a particular requirement from a customer that need to migrate Tenant2Tenant and wants to create migration batches.



Param(
  [bool]$silent=$false
)




# To run the script as a scheduled task or unattend script uncomment the following line
# $silent=$true

#$a=get-credential

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $a -Authentication Basic -AllowRedirection
#Import-PSSession $Session

# Get the list of mailboxes with required attribute
get-mailbox * -resultsize unlimited | select Identity,Name,DisplayName,samaccountname,RecipientType,RecipientTypeDetails,WindowsEmailAddress | export-csv mailboxes.csv -encoding "unicode" -NoTypeInformation


# Get the list of the groups, only the MailUniversalSecurityGroup (Exchange supported groups to provide permissions) are fetched
Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Select Name,Alias,DisplayName,WindowsEmailAddress,SamAccountName | export-csv groups.csv -NoTypeInformation

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
	$asd=Get-DistributionGroupMember $group 
	foreach ($user in $asd) {
			$row = $table.NewRow()
			$row.SamAccountNAme=$user.SamAccountname
			$row.DGName=$group 
			$row.DomainNB="Office365"
			$table.rows.add($row)
		}
	}
$table | export-csv groupsmembership.csv -encoding "unicode" -NoTypeInformation 
# End of Group Membership Block





# ADPermissions: Send-AS, Receive-AS
import-csv mailboxes.csv | Get-RecipientPermission | select Identity,Trustee,@{Name='AccessRight';Expression={"AD"}},ExtendedRight,FolderPath | export-csv ad.csv -encoding "unicode" -NoTypeInformation 

# Mailbox Permissions: FullAccess, Read
import-csv mailboxes.csv | Get-MailboxPermission | select Identity,User,@{Name='AccessRight';Expression={"MP"}},ExtendedRight,FolderPath | export-csv mp.csv -encoding "unicode" -NoTypeInformation 

# SendOnBeHalf permissions
$allSNB=import-csv mailboxes.csv | get-mailbox | ?{$_.GrantSendOnBehalfTo} 
"""Identity"",""User"",""AccessRight"",""ExtendedRight"",""FolderPath""" > SnB.csv 
foreach ($mailbox in $allSNB) {
	foreach ($SNB in $mailbox.GrantSendOnBehalfTo) {
		$id=$mailbox.Identity
		"$id,$SNB,SNB,," >> SnB.csv 
	}
}

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
        $folders = Get-MailboxFolderStatistics $identity.Identity | % {$_.folderpath} | % {$_.replace("/","\")}
	$alias=$identity.windowsEmailAddress
        $utente = $folders | %{ Get-MailboxFolderPermission $alias":"$_ | select Identity,User,Accessrights,foldername}


	foreach ($riga in $utente) {
		$acc=[string]::join(";",($riga.Accessrights))
		$row = $table.NewRow()
		$row.Identity=$identity.Identity
		$row.User=$riga.User
		$row.AccessRights="FP"
		$row.PermissionType="FP"
		$row.ExtendedRight=""
		$row.FolderPath="NoDetails"
		$table.rows.add($row)
                }
}

$table | export-csv fp.csv -encoding "unicode" -NoTypeInformation 
# End of get Folder Permissions block