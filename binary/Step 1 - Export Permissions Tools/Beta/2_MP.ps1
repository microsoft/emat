function ExportMPPerm {
    param (
   #     $FilePath = "c:\TEMP\",
        $FileName = "Mailboxes.csv",
        $ADPermsissionsFileName = "MP.csv",
        $PermissionFile ="MP_CommandTrace.csv" #File Permissions Trace
    )

$CSVFilePath = $FilePath+$FileName
$all = import-csv $CSVFilePath #-Delimiter ";"
Import-Module ActiveDirectory
#$initialLocation = (Get-Location).path
$FullNamePath= $FilePath+$ADPermsissionsFileName
$FullNamePathError = $FilePath+"ProcessError"+$ADPermsissionsFileName
Out-File -InputObject "`"Identity`",`"User`",`"AccessRight`",`"ExtendedRights`",`"FolderPath`"" -FilePath $FullNamePath -Encoding unicode
#Error File
#$FullNamePathError = $FilePath+"ProcessError"+$ADPermsissionsFileName


Set-ADServerSettings -ViewEntireForest $True

$Inc = 0
$all.count
$totalTimes = $all.count
$mbx = $Null
foreach ($mbx in $all ){
    #Write-Host "Working on: "$mbx.identity
    Write-Progress -Activity "Getting ACLs for $($mbx.identity)" -Status "Processing Mailbox $Inc of $totalTimes " -PercentComplete ($Inc / $all.count*100) #$percentComplete
    $MPs = $null
    $MPs = Get-MailboxPermission -id $mbx.identity | ? {$_.IsInherited -eq $False -and $_.AccessRights -match "FullAccess" -and $_.User -notmatch "NT AUTHORITY" -and $_.User -notmatch "S-1-5"}
    if ($MPs.AccessRights){
        foreach ($MP in $MPs) {
            #EMAT Output
            Out-File -InputObject "`"$($mbx.Identity)`",`"$($MP.user.rawidentity)`",`"MP`",," -FilePath $FullNamePath -Encoding unicode -Append
        }
    }
    $Inc = $Inc + 1
}
}

ExportMPPerm
