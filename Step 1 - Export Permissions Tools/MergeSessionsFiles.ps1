# Collect all sessions folders in the mail folder (where there is the big mailboxes.csv files) and execute this script
#




$folders=Get-ChildItem -Directory | where {$_.Name -like "Session*"}

foreach ($folder in $folders) {
	$ad+=import-csv $folder\ad.csv 
	$mp+=import-csv $folder\mp.csv 
	$fp+=import-csv $folder\fp.csv 
	$SnB+=import-csv $folder\SnB.csv

}

$ad | export-csv ad.csv -encoding "unicode" -NoTypeInformation
$mp | export-csv mp.csv -encoding "unicode" -NoTypeInformation
$fp | export-csv fp.csv -encoding "unicode" -NoTypeInformation
$SnB | export-csv SnB.csv -encoding "unicode" -NoTypeInformation