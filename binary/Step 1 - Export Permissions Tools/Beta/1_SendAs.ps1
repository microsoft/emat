#Export Recipient Information
function ExportMailboxes {
    param (
        $ExportFilePath = "mailboxes.csv"
    )
    try {
        Get-Mailbox -resultsize unlimited | select Identity,Name,DisplayName,samaccountname,RecipientType,RecipientTypeDetails,PrimarySmtpAddress | export-csv $ExportFilePath -encoding "unicode" -NoTypeInformation
     }
    catch {
        Write-Host "Impossible to get the mailboxes"
        brack
    }    
}

#Export AD Perminssions
function ExportADUserPerm {
    param (
    #    $FilePath = "c:\TEMP\",
        $FileName = "Mailboxes.csv",

        $ConfigurationPartitionDN = (Get-ADForest).PartitionsContainer.Replace("CN=Partitions,",""),
        $ADPermsissionsFileName = "AD.csv"
    )

# ADPermissions: Send-AS, Receive-AS
    $CSVFilePath = $FilePath+$FileName
    $identities = import-csv $CSVFilePath #-Delimiter ";"
    try{
    Import-Module ActiveDirectory
    }
    catch{
    exit
    }
    $initialLocation = (Get-Location).path
    Set-ADServerSettings -ViewEntireForest $True
    Set-Location AD:

    $Inc = 0
    $identities.count
    $totalTimes = $identities.count
    $ConfigurationPartitionSendAs = "CN=Send-As,CN=Extended-Rights,"+$ConfigurationPartitionDN
    $ConfigurationPartitionReceiveAs = "CN=Send-As,CN=Extended-Rights,"+$ConfigurationPartitionDN #Change to Receive-AS
    $SendAsGUID = (Get-ADObject -Identity $ConfigurationPartitionSendAs -Properties rightsGuid).rightsGuid
    $ReceiveAsGUID = (Get-ADObject -Identity $ConfigurationPartitionReceiveAs -Properties rightsGuid).rightsGuid
    $FullNamePath= $FilePath+$ADPermsissionsFileName
    $FullNamePathError = $FilePath+"ProcessError"+$ADPermsissionsFileName
    Out-File -InputObject "`"Identity`",`"User`",`"AccessRight`",`"ExtendedRights`",`"FolderPath`"" -FilePath $FullNamePath -Encoding unicode
    Out-File -InputObject "Identity,SamAccountName,ErrorMessage" -FilePath $FullNamePathError -Encoding unicode 

    
    foreach ($identity in $identities) {  
    #$percentComplete = ($Inc / $totalTimes) * 100
    Write-Progress -Activity "Getting ACLs for $($identity.SamAccountName)" -Status "Processing Mailbox $Inc of $totalTimes " -PercentComplete ($Inc / $identities.count*100) #$percentComplete
        try{
            try{
                # Exclusion of IsInherited, Only ActiveDirectoryRights type , Account Admin, Account System, Account Self, Account not S-1, Account match Account
                $AdACLs = (Get-Acl (Get-User $identity.identity).DistinguishedName -ErrorAction stop).Access | ? {($_.IsInherited -eq $False -and $_.ActiveDirectoryRights -eq "ExtendedRight") -and ` 
                    $_.IdentityReference -notmatch "Admin" -and $_.IdentityReference -notmatch $identity.SamAccountName -and  $_.IdentityReference -notmatch "Self" -and $_.IdentityReference -notmatch "S-1" -and `
                    $_.IdentityReference -notmatch "System" -and ($_.objectType -eq $SendAsGUID -or $_.objectType -eq $ReceiveAsGUID)}
            }
            catch{
                $AdACLs = (Get-Acl (Get-User $identity.SamAccountName).DistinguishedName -ErrorAction stop).Access | ? {($_.IsInherited -eq $False -and $_.ActiveDirectoryRights -eq "ExtendedRight") -and ` 
                    $_.IdentityReference -notmatch "Admin" -and $_.IdentityReference -notmatch $identity.SamAccountName -and $_.IdentityReference -notmatch "Self" -and $_.IdentityReference -notmatch "S-1" -and `
                    $_.IdentityReference -notmatch "System" -and ($_.objectType -eq $SendAsGUID -or $_.objectType -eq $ReceiveAsGUID)}
            }
            
            foreach ($AdAcl in $AdACLs){
                if($AdAcl.ObjectType -eq $SendAsGUID){
                    Out-File -InputObject "`"$($identity.identity)`",`"$($AdAcl.IdentityReference)`",`"Send-As`",`"AD`"," -FilePath $FullNamePath -Append -Encoding unicode
                }
                if($AdAcl.ObjectType -eq $ReceiveAsGUID){
                    #Removed
                    #Out-File -InputObject "`"$($identity.identity)`",`"$($AdAcl.IdentityReference)`",`"AD`",`"Receive-As`"," -FilePath $FullNamePath -Append -Encoding unicode
                }
            }
            $Inc = $Inc + 1
        }
        catch{
            #OLD $ACLs = Get-ADPermission $identity.Identity | where {$_.ExtendedRights.RawIdentity -eq "Send-As" -or $_.ExtendedRights.RawIdentity -eq "Receive-As"}
            try{
            <# Rimosso perchè include Receive-As 
            $ACLs = Get-ADPermission $identity.Identity | ? {$_.isInherited -eq $false -and $_.User -notmatch $identity.SamAccountName -and $_.User -notmatch "S-1" -and $_.User -notmatch "SELF" -and ` 
                $_.User -notmatch "Admin" -and ($_.ExtendedRights.RawIdentity -eq "Send-As" -or $_.ExtendedRights.RawIdentity -eq "Receive-As")} 
            #>
            $ACLs = Get-ADPermission $identity.Identity | ? {$_.isInherited -eq $false -and $_.User -notmatch $identity.SamAccountName -and $_.User -notmatch "S-1" -and $_.User -notmatch "SELF" -and ` 
                $_.User -notmatch "Admin" -and ($_.ExtendedRights.RawIdentity -eq "Send-As")}
            Foreach ($ACL in $ACLs){
                    #$acl | select Identity,User,@{Name='AccessRight';Expression={"AD"}},ExtendedRights,FolderPath | export-csv c:\temp\ad.csv -encoding "unicode" -NoTypeInformation -Append
                    Out-File -InputObject "`"$($ACL.Identity)`",`"$($ACL.User)`",`"AD`",`"$($ACL.ExtendedRights)`",`"$($ACL.FolderPath)`",`"`"" -FilePath $FullNamePath -Encoding unicode -Append
                }
            }
            catch{
                Out-File -InputObject "$($Identity.Identity),$($identity.SamAccountName),Error while exporting AD" -FilePath $FullNamePathError -Encoding unicode -Append
            }
            $Inc = $Inc + 1
        }
    }
    Set-Location $initialLocation
}

#   Conversion AD Permission  
#   Multiple Forest Environment
function SendAsTrace {
    param (
        $FilePath = "C:\Temp\",
        $MasterFile = "AD.csv",
        $PermissionFile ="EXO_Permissions_Mailbox.csv"
    )
    $ADACLs = import-csv $FilePath$MasterFile #-delimiter ";"
    $ADExportedACLs = $ADACLs | ? {$_.ExtendedRights -ne "Receive-As"}

    Out-File -InputObject "MailboxEmail;UserMailEmailaddress;AccessRight;ExtendedRights" -FilePath $FilePath$PermissionFile -Encoding unicode
    Foreach ($ADACL in $ADExportedACLs){
    $MailboxEmail = $null
    $UserMail = $null
	try{
		$MailboxEmail = (Get-Recipient $ADACL.Identity).PrimarySmtpAddress.Address
	}
	Catch{
		Write-Host "Impossible to Get Recipient"
		break
	}
	if ($MailboxEmail){
		try {
			try {
			#Getting email address AD User from Account Forest 
				$UserMail = (Get-ADUser ($ADACL.User -split "\\","")[1] -Properties emailaddress -Server ($ADACL.User -split "\\","")[0]).emailaddress
				}
			catch{
			#Getting Email Address from Resource Forest 
				#$UserMail = (Get-ADUser $(($ADACL.User -split "\\","")[1] -replace "MIG_","") -Properties emailaddress ).emailaddress
                #Group Perm
                $UserMail = (Get-ADGroup ($ADACL.User -split "\\","")[1] -Properties mail -Server ($ADACL.User -split "\\","")[0]).mail
			}
			
		}
		catch{
		#Error On getting Get Ad User from Account and Resource forest, forcing SAM account name ad Trustee
			$UserMail = $ADACL.User
		}
		if ($null -eq $UserMail){
		#No Error Generated from Try and Catch and AD User Email address is empty, forcing SAM account name ad Trustee
		$UserMail = $ADACL.User
		}
	}
	Out-File -InputObject "$MailboxEmail;$($UserMail);ExtendedRight;$($ADACL.AccessRight)" -FilePath $FilePath$PermissionFile -Encoding unicode -append
    }
}

ExportMailboxes

ExportADUserPerm

SendAsTrace
