<# Step 1 - Download EWS Managed API
 As of 30th November 2020, I found it here:
 https://www.microsoft.com/en-ca/download/details.aspx?id=42951

 BUT the best way to find it now is through www.nuget.org as it's a bit more recent
 https://www.nuget.org/packages/Exchange.WebServices.Managed.Api/
 to install it, first register www.nuget.org in your machine:

 Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Force

 #>

Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Force
<#
Name                             ProviderName     IsTrusted  Location
----                             ------------     ---------  --------
nuget.org                        NuGet            False      https://www.nuget.org/api/v2    
#>

Get-PackageSource
<#
Name                             ProviderName     IsTrusted  Location
----                             ------------     ---------  --------
nuget.org                        NuGet            False      https://www.nuget.org/api/v2
PSGallery                        PowerShellGet    False      https://www.powershellgallery.com/api/v2
#>

Install-Package Exchange.WebServices.Managed.Api

Get-Package *Exchange* |ft -a
<#
Name                                            Version     Source
----                                            -------     ------
Exchange.WebServices.Managed.Api                2.2.1.2     C:\Program Files\PackageManagement\NuGe...
Microsoft Exchange Web Services Managed API 2.2 15.0.913.18
ExchangeOnlineManagement                        2.0.3       https://www.powershellgallery.com/api/v2
#>

# Step 2 - load the EWS API
$dllpath = "C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

# Step 3 - Create an EWS Service Object
$exchangeservice.UseDefaultCredentials = $true
$exchangeservice.AutodiscoverUrl($mailbox)

# Step 4 - Bind to the Inbox
$inboxfolderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox)
$inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$inboxfolderid)

# Step 5 - Configure Search Filter
$sfunread = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$sfsubject = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring ([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $subjectfilter)
$sfattachment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfcollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfcollection.add($sfunread)
$sfcollection.add($sfsubject)
$sfcollection.add($sfattachment)

# Step 6 - View results
$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 10
$foundemails = $inboxfolder.FindItems($sfcollection,$view)

<#
Then I just call FindTargetFolder($processedfolderpath) and FindTargetEmail($subject) and you're done.
#>
