# ProvisionningSharepointO365
## Setting up an Azure AD app

### Delegated Permissions (acting in the name of the user)

In this section you can learn how to register an application in Azure Active Directory.

#### Configuring the application in Azure AD

In this step by step guide you will register an application in Azure Active Directory, in order to consume PnP Framework in the name of the user connected to your app (i.e. with a delegated access token). Follow below steps to configure an application in Azure AD:

1. Navigate to https://aad.portal.azure.com/
2. Click on **Azure Active Directory** from the left navigation
3. Click on **App registrations** in the **Manage** left navigation group
4. Click on **New registration**
5. Give the application a name (e.g. PnP Framework) and click on **Register**
6. Copy the **Application ID** (Client ID) from the **Overview** page, you'll need this GUID value later on
7. Copy the **Directory ID** (Tenant ID) from the **Overview** page, you'll need this GUID value later on
8. Click on the **API Permissions** in the **Manage** left navigation group
9. Click on **Add Permissions** and add the permissions you want to give to this application. Below list is a recommendation, you can grant less permissions but that might result in some calls to fail due getting access denied errors.

   - SharePoint -> Delegated Permissions -> AllSites -> AllSites.FullControl
   - SharePoint -> Delegated Permissions -> Sites -> Sites.Search.All
   - SharePoint -> Delegated Permissions -> TermStore -> TermStore.ReadWrite.All
   - SharePoint -> Delegated Permissions -> User -> User.ReadWrite.All
   - Microsoft Graph -> Delegated Permissions -> User -> User.Read
   - Microsoft Graph -> Delegated Permissions -> Directory -> Directory.ReadWrite.All
   - Microsoft Graph -> Delegated Permissions -> Directory -> Directory.AccessAsUser.All
   - Microsoft Graph -> Delegated Permissions -> Group -> Group.ReadWrite.All

10. Click on the **Grant admin consent for** button to consent to these permissions for the users in your organization
11. Click on **Authentication** in the **Manage** left navigation group
12. Change **Default client type** to **Treat application as public client** and hit **Save** (this step is optional and you should do that if and only if you are planning to use username + password for authentication)

If you want to configure support for interactive login you should also configure the _Platform_ and the _redirect URI_ in the **Authentication** panel. You can read [further details here](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app#add-a-redirect-uri).

13. Click on **Authentication** and then click on **Add a platform**, choose **Mobile and desktop applications** and provide http://localhost as the **Redirect URI**

### Application Permissions (acting as an app account with app-only permissions)

In this section you can learn how to register an application in Azure Active Directory and how to use it in your .NET code, in order to use the PnP Framework within a background job/service/function, running your requests with an app account.

#### Configuring the application in Azure AD

The easiest way to register an application in Azure Active Directory for app-only is to use the [PnP PowerShell](https://pnp.github.io/powershell) cmdlets. Specifically you can use the [`Register-PnPAzureADApp` command](https://github.com/pnp/powershell/blob/dev/documentation/Register-PnPAzureADApp.md) with the following syntax:

```powershell
$app = Register-PnPAzureADApp -ApplicationName "PnP.Framework.Consumer" -Tenant contoso.onmicrosoft.com -OutPath c:\temp -CertificatePassword (ConvertTo-SecureString -String "password" -AsPlainText -Force)  -GraphApplicationPermissions "Group.ReadWrite.All", "User.ReadWrite.All" -SharePointApplicationPermissions "Sites.FullControl.All", "TermStore.ReadWrite.All", "User.ReadWrite.All"  -Store CurrentUser -DeviceLogin
```

The above command will register for you in Azure Active Directory an app with name `PnP.Framework.Consumer`, with a self-signed certificate that will be also saved on your filesystem under the `c:\temp` folder (remember to create the folder or to provide the path of an already existing folder), with a certificate password value of `password` (you should provide your own strong password, indeed). Remember to replace `contoso.onmicrosoft.com` with your Azure AD tenant name, which typically is `company.onmicrosoft.com`. The permissions granted to the app will be:

   - SharePoint -> Application Permissions -> Sites -> Sites.FullControl.All
   - SharePoint -> Application Permissions -> TermStore -> TermStore.ReadWrite.All
   - SharePoint -> Application Permissions -> User -> User.ReadWrite.All
   - Microsoft Graph -> Application Permissions -> User -> User.ReadWrite.All
   - Microsoft Graph -> Application Permissions -> Group -> Group.ReadWrite.All

Executing the command you will first have to authenticate against the target tenant, providing the credentials of a Global Tenant Admin. Then you will see a message like the following one:

```text
Waiting 60 seconds to launch consent flow in a browser window. This wait is required to make sure that Azure AD is able to initialize all required artifacts.........
```

Almost 60 seconds later, the command will prompt you for authentication again and to grant the selected permissions to the app you are registering. Once you have done that, in the `$app` variable you will find information about the just registered app. You can copy in your clipboard the **Application ID** (Client ID) executing the following command:

```powershell
$app.AzureAppId | clip
```

And you can copy in your clipboard the thumbprint of the generated X.509 certificate executing the following command:

```powershell
$app.'Certificate Thumbprint' | clip
```

Paste this copied values in a safe place, because you will need them later on to setup authentication. In the `c:\temp` folder (or whatever else folder you will choose) there will also be a file named `PnP.Framework.Consumer.pfx`, which includes the private key of the self-signed certificate generated for you, as well as a file named `PnP.Framework.Consumer.cer`, which includes the public key of the self-signed certificate generated for you.

## What are the typical code changes I need to make?

PnP Framework is for 90% identical to PnP Sites Core, we did drop some legacy components and everything that was specific for on-premises but most folks will not be impacted by that. The main change that will require changes in your code is due to authentication: the underlying CSOM for .NET Standard library only supports access token based authentication so we had to refactor the AuthenticationManager class so that we use Microsoft Azure AD based OAuth to authenticate.

We unified authentication between PnP Framework, PnP PowerShell and PnP Core SDK to use [Microsoft.Identity.Client (MSAL)](https://github.com/AzureAD/ microsoft-authentication-library-for-dotnet) as underlying model. Whereas in PnP Sites Core you would provide the needed auth information when invoking an AuthenticationManager method call you now pass that information via the constructor and then use a generic method to request an access token:

### Delegated authentication
