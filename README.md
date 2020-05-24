# azure-devops-powershell
 
 Simple demostration to call Azure DevOps Rest APIs with Powershell.
 This demo creates;
 - Azure AD Application
 - Service Principal
 - Azure DevOps Project
 - Azure DevOps Project's Service Connection
 
```powershell
./New-AzDevOpsProject.ps1 -TenantId "11111111-2222-3333-4444-555555555555" `
 -OrganizationName "SomeOrganization" `
 -SubscriptionId "99999999-8888-7777-6666-555555555555" `
 -PAT "123abc456" `
 -ProjectName "TestProject" `
 ```

Azure DevOps Rest API: https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-6.0
