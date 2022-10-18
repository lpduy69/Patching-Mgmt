$Criteria = "IsInstalled=0 and Type='Software'"
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$SearchResult = $Searcher.Search($Criteria).Updates
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Operation Completed",0,"Done",4096)
$count = ($SearchResult | Measure-Object).count
If ($count -eq 0) 
{
	Write-Host "Nothing found, abnormal except if this is second check, need to restart server and launch again" -ForegroundColor Red
}else{
	$SearchResult | Select-Object { $_.Title }
}

