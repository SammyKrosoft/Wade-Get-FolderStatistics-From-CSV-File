# First populate a users.csv file with at least a column with user names on it, if you don't have one handy, you can use a command line like dis:
# get-mailbox -Filter {Recipienttypedetails -eq "UserMailbox"} |Select name, primarysmtpaddress |Export-Csv -NoTypeInformation c:\temp\users.csv

$StopWatch = [System.diagnostics.stopwatch]::StartNew()
$OutputFolder = "C:\temp\"
$OutputCSV = $OutputFolder + "MailboxFolderStats_" + (Get-Date -Format "MMddyyyy_HHmmss") + ".csv"
$AllUsers = Import-csv users.csv

$TotalUsers = $AllUsers.Count
$UserFolderStatsCollection = @()
$Counter = 0
Foreach ($User in $AllUsers) {
    Write-Progress -Activity "Collecting Folders stats" -Status "Processing User $($User.Name)" -PercentComplete ($Counter/$TotalUsers*100)
    $Counter++
    $UserFolderStatsCollection += Get-MailboxFolderStatistics -id $User.Name | Select FolderPath, FolderSize, ItemsInFolder | where {$_.FolderSize -ne 0} | Foreach {Add-Member -InputObject $_ -Name MailboxName -MemberType NoteProperty -Value $User.Name -PassThru ;  Add-Member -InputObject $_ -Name FolderSizeInBytes -MemberType NoteProperty -Value ($_.FolderSize.ToBytes()); Add-Member -InputObject $_ -Name FolderSizeInMB -MemberType NoteProperty -Value ($_.FolderSize.ToMB())}
}

$UserFolderStatsCollection | Select MailboxName,FolderPath,FolderSize,FolderSizeInBytes,FolderSizeInMB,ItemsInFolder | Export-CSV -NoTypeInformation $OutputCSV
$StopWatch.Stop()
$ScriptDuration = [math]::Round($StopWatch.Elapsed.TotalSeconds,1)

Write-host "Script took $ScriptDuration seconds to execute..."

notepad $OutputCSV 
