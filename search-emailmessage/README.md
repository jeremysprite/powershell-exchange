# Search-EmailMessage

Search for and delete an email using PowerShell and the EWS API

## Requirements:
Install the EWS Managed API https://www.microsoft.com/en-us/download/details.aspx?id=42951

Set up Impersonation Rights
https://msdn.microsoft.com/en-us/library/office/dn722376(v=exchg.150).aspx

## Usage
```powershell
Search-EmailMessage.ps1 -MailboxToImpersonate "user@sysadminasaservice.com" -AccountWithImpersonationRights "audit-mailbox@sysadminasaservice.com"
  -searchString "Hello World" -delete $false
```  
