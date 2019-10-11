# Create an Azure AD Application

## Choose the tenant where you want to create your app

1. Sign in to the [Azure portal](https://portal.azure.com) using either a work or school account.
1. If your account is present in more than one Azure AD tenant:
   1. Select your profile from the menu on the top right corner of the page, and then **Switch directory**.
   1. Change your session to the Azure AD tenant where you want to create your application.

## Register the app

1. Navigate to the [Azure portal > Azure Active Directory > App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) to register your app.
2. Select **New registration**.

![Add a new application registration](./media/setup01_AppReg.png 'Perform a Search')

3. When the **Register an application page** appears, enter your app's registration information:
   1. In the **Name** section, enter a meaningful name that will be displayed to users of the app. For example: `GraphSearchApi`
   1. In the **Supported account types** section, select **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.
4. Select **Register** to create the app.

![Add a new application registration](./media/setup02_NewAppReg.png 'Perform a Search')

5. On the app's **Overview** page, find the **Application (client) ID** value and record it for later. You'll need this value to configure the Visual Studio configuration file for this project.

6. In the list of pages for the app, select **Authentication**.

7. In the **Redirect URIs** section, select **Web** in the combo-box and enter the following redirect URIs:

- `http://localhost:44308/`
- `http://localhost:44308/signin-oidc`

> NOTE: All the Microsoft Graph Search API samples in this repo are designed to run on port 44308.  If you create your own solution, be sure to modify this setting.

8. In the **Advanced settings** > **Implicit grant** section, check **ID tokens** as Sample 1.0 requires the [Implicit grant flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-implicit-grant-flow) to be enabled to sign-in the user and call an API.

![Set application authentication settings](./media/setup03_AppAuthSettings.png 'Setup auth settings')

9. Select **Save**.

10. From the **Certificates & secrets** page, in the **Client secrets** section, choose **New client secret**.
   
1. Enter a key description (of instance `app secret`).

1. Select a key duration of either **In 1 year**, **In 2 years**, or **Never Expires**.

1. When you click the **Add** button, the key value will be displayed. Copy the key value and save it in a safe location.

>**NOTE** You'll need this key later to configure the project in Visual Studio. This key value will not be displayed again, nor retrievable by any other means, so record it as soon as it is visible from the Azure portal.

![Create new app secret](./media/setup04_AppSecret.png 'Create App Secret')

1. In the list of pages for the app, select **API permissions**.

1. Click the **Add a permission** button and then make sure that the **Microsoft APIs** tab is selected.

![Add app permissions](./media/setup05_AddPermissions.png 'Add App Permissions')

1. In the **Commonly used Microsoft APIs** section, select **Microsoft Graph**.

1. In the **Delegated permissions** section, make sure that the following *delegated* permissions are checked: 

   -  **Calendars.Read**
   -  **Files.Read.All**
   -  **Mail.Read**
   -  **User.Read**
   -  **email**
   -  **office_access**
   -  **openid**
   -  **profile**

![Add graph permissions](./media/setup06_MSGraphPermissions.png 'Add Graph Permissions')

>**NOTE** These permissions will allow the sample application(s) to read data from the Microsoft Graph and retrieve information about users from Azure Active Directory via the Microsoft Graph API.

![Final graph permissions](./media/setup06_MSGraphPermissionsFinal.png 'Final Graph Permissions')

14. Select the **Add permissions** button.

## Grant Admin consent to view Security data

### Assign Scope (permission)

1. If you are an Azure Admin, click the **Grant admin consent for YOURTENANT** button

![Grant consent](./media/setup08_GrantConsent.png 'Grant Consent')

1. If you are not an Azure AD Administrator, provide your administrator the **Application Id** and the **Redirect URI** that you used in the previous steps. The organization’s Azure Active Directory Tenant Administrator is required to grant the required consent (permissions) to the application.

2.	As the Tenant Administrator for your organization, open a browser window and paste the following URL in the address bar (after replacing the values for TENANT_ID, APPLICATION_ID and REDIRECT_URL):
https://login.microsoftonline.com/TENANT_ID/adminconsent?client_id=APPLICATION_ID&state=12345&redirect_uri=REDIRECT_URL.

> **Note:** Tenant_ID is the same as the AAD Directory ID, which can be found in the Azure Active Directory Blade within [Azure Portal](portal.azure.com). To find your directory ID, Log into [Azure Portal](portal.azure.com) with a tenant admin account. Navigate to “Azure Active Directory”, then “Properties”. Copy your ID under the "Directory ID" field to be used as **TENANT_ID**.

3.	After authenticating, the Tenant Administrator will be presented with a dialog like the following (depending on the permissions the application is requesting)

![Grant admin consent](./media/setup09_GrantConsentAdmin.png 'Grant Admin Consent')

4. By clicking on "Accept" in this dialog, the Tenant Administrator is granting consent to all users of this organization to use this application. Now this application will have the correct scopes (permissions) need to access the Security API, the next section explains how to authorize a specific user within your organization (tenant).

>**Note:** Because there is no application currently running at the redirect URL you will be receive an error message. This behavior is expected. The Tenant Administrator consent will have been granted by the time this error page is shown.

### Authorize users in your organization to access the Microsoft Graph security API

To access security data through the Microsoft Graph security API, the client application must be granted the required permissions and when operating in Delegated Mode, the user signed in to the application must also be authorized to call the Microsoft Graph security API.

This section describes how the Tenant Administrator can authorize specific users in the organization.

1. As a Tenant **Administrator**, sign in to the [Azure Portal](https://portal.azure.com).

2. Navigate to the Azure Active Directory blade.

3. Select **Users**.

4. Select a user account that you want to authorize to access to the Microsoft Graph security API.

5. Select **Directory Role**.

6. Select the **Limited Administrator** radio button and select the check box next to **Security administrator** role

7. Click the **Save** button at the top of the page

Repeat this action for each user in the organization that is authorized to use applications that call the Microsoft Graph security API. Currently, this permission cannot be granted to security groups.

> **Note:** For more details about the authorization flow, read [Authorization and the Microsoft Graph Security API](https://developer.microsoft.com/en-us/graph/docs/concepts/security-authorization).
