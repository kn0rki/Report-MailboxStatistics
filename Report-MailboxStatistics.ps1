param (
  $users
)

$CountTopFolder = 10
$ReportMailboxSizeInMB = 1800

$SMTPServer = ""
$From = ""
$Subject = "Postfach Übersicht"

[System.Collections.ArrayList]$MailboxStatistics = @()
foreach ($user in $users) {
 $Mailbox = get-mailbox $user.SamAccountName
 $EMail = $Mailbox.PrimarySmtpAddress
 $Stats = $Mailbox | Get-MailboxStatistics | Select-Object displayname, @{label="Size"; expression={$_.TotalItemSize.Value.ToMB()}}
 $Displayname = $Stats.Displayname
 $MailboxSize = $Stats.Size
 if ($MailboxSize -ge $ReportMailboxSizeInMB) {
  $MailboxFolderStatistics = Get-MailboxFolderStatistics $mailbox | Select-Object FolderPath,FolderSize,ItemsInFolder

  $TopFoldersBySize = $MailboxFolderStatistics |
    Select-Object FolderPath,@{
      Name="Foldersize";Expression={
        [long]$a = "{0:N2}" -f ((($_.FolderSize -replace "[0-9\.]+ [A-Z]* \(([0-9,]+) bytes\)","`$1") -replace ",","")); [math]::Round($a/1MB,2) }}  |
    Sort-Object foldersize -Descending |
    Select-Object -first $CountTopFolder

  $TopFoldersByItems = $MailboxFolderStatistics | Sort-Object ItemsInFolder -Descending | Select-Object -first $CountTopFolder

  $Statistic = [PSCustomObject]@{
	 DisplayName = $Displayname
	 EMail = $EMail
	 MailboxSize = $MailboxSize
	 TopFoldersBySize = $TopFoldersBySize
	 TopFoldersByItems = $TopFoldersByItems
	}
  $null = $MailboxStatistics.Add($Statistic)
 }
}

foreach ($MailboxStatistic in $MailboxStatistics) {
  $MailboxSize = $MailboxStatistic.MailboxSize
  $TopFoldersBySize = $MailboxStatistic.TopFoldersBySize |
    Select-Object @{
      label="Ordnerpfad"; expression={$_.Folderpath}
    }, @{
        label="Größe"; expression={$str = $_.Foldersize; [string]$str + " MB"}
      } | ConvertTo-Html -Fragment

  $TopFoldersByItems = $MailboxStatistic.TopFoldersByItems |
    Select-Object @{
      label="Ordnerpfad"; expression={$_.Folderpath}
    },
    @{
      label="Anzahl Elemente"; expression={$_.ItemsInFolder}
    } | ConvertTo-Html -Fragment
  $To =  $MailboxStatistic.EMail
 $MailBody = @"
 <!DOCTYPE html>
 <html lang="de">
  <head>
   <title>Mailbox Report</title>
   <style>
    body {font-family: Calibri;}
    td {width:100px; max-width:300px; background-color:white;}
    table {width:100%;}
    th {text-align:left; font-size:12pt; background-color:lightgrey;}
   </style>
  </head>
 <body>
  <h2>Mailbox Übersicht</h2>
 <div><p>Ihr Postfach ist $MailboxSize MB groß, bitte löschen Sie nicht mehr benötigte Daten aus Ihrem Postfach.</p></div>
 <div><p>Dies ist eine Übersicht ihrer $CountTopFolder größten Ordner in ihrem Postfach:</p></div>
 $TopFoldersBySize
 <div><p>Ordner mit vielen Elementen beeinträchtigen die Outlook Geschwindigkeit, löschen Sie nicht mehr benötigte Elemente um Outlook nicht zu verlangsamen.<br/>Dies sind Ihre $CountTopFolder Ordner mit den meisten Elementen:</p></div>
 $TopFoldersByItems
 </body></html>
"@

 Send-MailMessage -SmtpServer $SMTPServer -From $From -To $To -Body $MailBody -BodyAsHtml -Encoding UTF8 -Subject $Subject
}
