param(
    [string]$MailboxToImpersonate = "user1@sysadminasaservice.com",
    [string]$AccountWithImpersonationRights = "audit-mailbox@sysadminasaservice.com",
    [System.Net.NetworkCredential]$creds = (Get-Credential),
    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010,
    [string]$searchString = "Hello World",
    [bool]$delete = $false
    
)


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

# Use below if you need to hardcode credentials into the script (for example, for unattended script running)
# Also, it is better to use WebCredentials instead of NetworkCredentials like the below - https://blogs.msdn.microsoft.com/webdav_101/2015/05/11/best-practices-ews-authentication-and-access-issues/
<#
$secpassword = ConvertTo-SecureString "VMuue2wZ4aqgcxkX1sf6" -AsPlainText -Force
$psCred = New-Object System.Management.Automation.PSCredential("audit-mailbox@sysadminasaservice.com",$secpassword)
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
#>

# Instantiate Exchange Service object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

# Set the Service credentials
$service.Credentials = $creds

# Set the URL of the CAS (Client Access Server) - either use Autodiscover or hardcoded URL if Autodiscover not set up
$service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
# $service.Url = new-object Uri("https://owa.sysadminasaservice.com/EWS/Exchange.asmx");

# Login to Mailbox with Impersonation
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate );

# Connect to the Inbox
$InboxFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName)
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$InboxFolder)

# Search for an email(s)
# Below we are searching for an unread email with the subject equal to our searchString parameter
$SearchFilterCollection = @()
$SearchFilter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,$searchString)
$SearchFilter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$SearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$SearchFilterCollection.Add($SearchFilter1)
$SearchFilterCollection.Add($SearchFilter2)

# Create an item view, basically how many results you want to display, and a property set
$iv = new-object Microsoft.Exchange.WebServices.Data.ItemView(50)
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 

$messages = $service.FindItems($Inbox.Id, $SearchFilterCollection, $iv)

# Now, we are going to work with each email message
If($messages.TotalCount -gt 0){
    foreach($message in $messages){
       
        Write-Host "Found message! Message ID is $($message.Id.UniqueId)"
        if($delete -eq $true){
            # HARD delete message - can also SoftDelete (into the Dumpster). MoveToDeletedItems is done by a Move request.
            $emailMessage = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Message.Id.UniqueId,$propertySet)
            $emailMessage.Delete("HardDelete")
        }

    }

}else{

    Write-Host "Search returned no results"

}