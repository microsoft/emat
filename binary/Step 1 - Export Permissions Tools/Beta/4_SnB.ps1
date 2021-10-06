
function ExportSnBPerm {
    param (
 #       $FilePath = "c:\TEMP\",
        $FileName = "Mailboxes.csv",
        $ADPermsissionsFileName = "SnB.csv",
        $PermissionFile ="SnB_CommandTrace.csv" #File Permissions Trace
    )

$CSVFilePath = $FilePath+$FileName
$all = import-csv $CSVFilePath #-Delimiter ";"
Import-Module ActiveDirectory
#$initialLocation = (Get-Location).path
$FullNamePath= $FilePath+$ADPermsissionsFileName
$FullNamePathError = $FilePath+"ProcessError"+$ADPermsissionsFileName
Out-File -InputObject "`"Identity`",`"User`",`"AccessRight`",`"ExtendedRights`",`"FolderPath`"" -FilePath $FullNamePath -Encoding unicode
#File Permissions Trace
Out-File -InputObject "`"Identity`",`"User`",`"AccessRight`",`"ExtendedRights`",`"FolderPath`""  -FilePath $FilePath$PermissionFile -Encoding unicode
#Error File
$FullNamePathError = $FilePath+"ProcessError"+$ADPermsissionsFileName
Out-File -InputObject "Identity,User,ErrorMessage" -FilePath $FullNamePathError -Encoding unicode 

Set-ADServerSettings -ViewEntireForest $True
$mbx = $Null
foreach ($mbx in $all ){
    Write-Host "Working on: "$mbx.identity
    $SnBs = $null
    $SnBs = Get-mailbox -id $mbx.identity #| ? {$_.GrantSendOnBehalfTo} 
    if ($SnBs.GrantSendOnBehalfTo){
        foreach ($SnB in $SnBs.GrantSendOnBehalfTo) {
            #Emat Output
            Out-File -InputObject "`"$($mbx.Identity)`",`"$($SnB)`",`"SnB`",," -FilePath $FullNamePath -Encoding unicode -Append
            try {
                $User = (get-Recipient $SnB -erroraction stop).PrimarySmtpAddress #taking Mailbox and Groups
                Out-File -InputObject "`"$($SnBs.PrimarySmtpAddress)`",`"$($User)`",`"SnB`",," -FilePath $FilePath$PermissionFile -Encoding unicode -Append
            }
            catch {
                Write-Host "no mailbox "$SnB -ForegroundColor Red
                Out-File -InputObject "$($mbx.Identity),$($SnB),User Not found as recipient" -FilePath $FullNamePathError -Encoding unicode -Append
            }
            

        }
    }
}
}

ExportSnBPerm
