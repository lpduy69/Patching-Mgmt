$Session = New-Object -ComObject "Microsoft.Update.Session"
$Searcher = $Session.CreateUpdateSearcher()
$historyCount = $Searcher.GetTotalHistoryCount()
$fixes = $Searcher.QueryHistory(0, $historyCount) | Select-Object Date, Title, operation, @{name="Status"; expression={switch($_.resultcode){ 1 {"In Progress"}; 2 {"Succeeded"}; 3 {"Succeeded With Errors"}; 4 {"Failed"}; 5 {"Aborted"}}}} | Where-Object {$_.Date -ge (Get-Date).AddDays(-1) -And $_.operation -eq 1} | Sort-Object Date | Select-Object Title, Date, Status | Format-List *
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Operation Completed",0,"Done",4096)
$fixes
