
# Get the list of the groups, only the MailUniversalSecurityGroup (Exchange supported groups to provide permissions) are fetched
Get-Group -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Select Name,Alias,DisplayName,WindowsEmailAddress,SamAccountName | export-csv groups.csv -encoding "unicode" -NoTypeInformation

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
