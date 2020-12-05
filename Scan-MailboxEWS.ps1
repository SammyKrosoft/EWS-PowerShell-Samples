param($mailboxName = "new-tickets@contoso.com",
$smtpServerName = "smtp.contoso.com",
$emailFrom = "monitorservice@contoso.com",
$emailTo = "support@contoso.com"
)
 
# Load the EWS Managed API
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
 
try {
  $Exchange2007SP1 = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
  $Exchange2010    = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010
  $Exchange2010SP1 = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
  $Exchange2010SP2 = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
  $Exchange2013    = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
  $Exchange2013SP1 = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
 
  # create EWS Service object for the target mailbox name
  $exchangeService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $Exchange2010SP2
  $exchangeService.UseDefaultCredentials = $true
  $exchangeService.AutodiscoverUrl($mailboxName)
 
  # bind to the Inbox folder of the target mailbox
  $inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
  $inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
 
  # Optional: reduce the query overhead by viewing the inbox 10 items at a time
  $itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 10
  # search the mailbox for messages older than 15 minutes
  $dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
  $15MinutesAgo = (Get-Date).AddMinutes(-15)
  $searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$15MinutesAgo
  $foundItems = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView)
 
  # report the results via email and Application event log
  $entryType = "Information"
  $messageBody = "Self-service mailbox scan completed at {0}.`r`n" -f (get-date -format "MM/dd/yyyy hh:mm:ss")
  if ($foundItems.TotalCount -ne 0) {
  $entryType = "Warning"
  $subject = "Self-service mailbox hung"
  $messageBody  = "Inbox has {0} message(s) that are more than 15 minutes old.`r`n" -f $foundItems.TotalCount
  $messageBody += "Inbox has {0} message(s) total.`r`n`r`n" -f $inboxFolder.TotalCount
  $messageBody += "Please restart the Email Engine on SERVER01`r`n"
  $messageBody += "Self-service mailbox scan completed at {0}.`r`n" -f (get-date -format "MM/dd/yyyy hh:mm:ss")
  $messageBody += "Script run from $env:computername`r`n"
  $smtpClient = New-Object -TypeName Net.Mail.SmtpClient -ArgumentList $smtpServerName
  $smtpClient.Send($emailFrom, $emailTo, $subject, $messageBody)
  }
  Write-EventLog -LogName "Application" -Source "Application" -EventId 1 -Category 4 -EntryType $entryType -Message $messageBody
}
catch
{
  $entryType = "Error"
  $subject = "Error in mailbox monitor script"
  $messageBody = "{0}`r`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage
  Write-EventLog -LogName "Application" -Source "Application" -EventId 1 -Category 4 -EntryType $entryType -Message $messageBody
  $smtpClient = New-Object -TypeName Net.Mail.SmtpClient -ArgumentList $smtpServerName
  $smtpClient.Send($emailFrom, $emailTo, $subject, $messageBody)
}