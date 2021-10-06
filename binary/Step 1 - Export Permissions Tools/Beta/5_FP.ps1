
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
        $alias = $identity.PrimarySmtpAddress
        $folders = (Get-MailboxFolderStatistics $alias | ?{$_.FolderType -eq "calendar"}).FolderPath -replace ("/","\")
        $utente = $folders | %{ Get-MailboxFolderPermission $alias":"$_  -ErrorAction silentlycontinue | select Identity,User,Accessrights,foldername} | where {$_.User -notlike "Default*" -and $_.User -notlike "Anonymous*"}
	

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