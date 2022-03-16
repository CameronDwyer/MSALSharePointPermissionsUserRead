# MSALSharePointPermissionsUserRead

To run this sample:

Substitute your SharePoint tenant URL in the program.cs file

```
const string sharepointTenantUrl = "https://YOUR-TENANT.sharepoint.com"; // <- REPLACE YOUR SHAREPOINT TENANT HERE
```

The application registration is hosted in my tenant and you can use it, otherwise create an application registration in your own tenant using the Azure CLI and this command. Replace the resulting application id (client id) in the code (program.cs).

```
az ad app create --display-name MSALSharePointPermissions --native-app true --reply-urls http://localhost --available-to-other-tenants true
```

When prompted for authentication ensure you sign-in with an identity from the same tenant as the SharePoint URL you provided in code.