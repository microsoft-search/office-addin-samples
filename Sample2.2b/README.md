# Microsoft Graph Search API Sample for Excel (Node.js)

## Table of contents

* [Introduction](#introduction)
* [Prerequisites](#prerequisites)
* [Getting started with the sample](#getting-started-with-the-sample)
* [Build and run the sample](#build-and-run-the-sample)
* [Code of note](#code-of-note)
* [Questions and comments](#questions-and-comments)
* [Contributing](#contributing)
* [Additional resources](#additional-resources)

## Introduction

This sample includes a NodeJS/Express application that demonstrates how to add an AddIn to Outlook and trade and IdentityAPI token for a Graph token to make a call to Microsoft Graph Search API endpoint.

## Prerequisites

This sample requires the following:  

  * [Visual Studio Code](https://code.visualstudio.com/) 
  * [Outlook 2016 or higher](https://docs.microsoft.com/en-us/office365/troubleshoot/administration/switch-channel-for-office-365)
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with sample

 1. Download or clone this repo.

### Create an Azure AD Application

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Configure the project

1.  Open Visual Studio code to the ./Sample2.2a/SearchBetaApi folder
2.  Open the **server.ts** file
3.  Replace the values for the following:

- {client TenantID} - your Azure AD Tenant ID
- {client GUID}- the client id from the Configuring Azure steps
- {client secret} - the client secret from the Configuring Azure steps

>NOTE:  If you did the Visual Studio sample in **Sample 2.0**, you can skip to **Test the AddIn** as all the settings are the same.

## Update the AddIn manifest xml

1.  Open the **GraphSearchApiOutlook.xml** file
2.  Scroll to the bottom of the file, in the **WebApplicationInfo** section, ensure that the clientid matches the client id from the Configuring Azure steps
3.  Save the file

## Test the AddIn

1.  Switch back to Visual Studio Code
2.  Click the **debug** tab, then select **Launch Program** configuration
3.  Switch back to Excel, click **Retry** to refresh the application task pane.
4.  Run a search, review the results that are exported to the workbook sheet as a filterable and searchable table

## Code of note

- Remember if you make changes, you must run **npm run-script build** to rebuild the TypeScript files into their corresponding javascript.
- The **server.ts** file contains the endpoint that will trade the identity token for the graph token.

## Questions and comments

We'd love to get your feedback about this sample! 
Please send us your questions and suggestions in the [Issues](https://github.com/microsoftgraph/aspnet-connect-rest-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoftgraph).
Tag your questions with [MicrosoftGraph].

## Contributing ##

If you'd like to contribute to this sample, see [CONTRIBUTING.md](CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). 
For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Additional resources

- [Microsoft Graph Security API Documentaion](https://aka.ms/graphsecuritydocs)
- [Other Microsoft Graph Connect samples](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph overview](https://graph.microsoft.io)
- [Sideload Outlook add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)
- [Tutorial: Build a message compose Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial?context=office/dev/add-ins/context)
- [Deploy and publish your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish)
- [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/use-rest-api)
- [Identity API requirement sets](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)
- [Enable single sign-on for Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins#configure-the-add-in)
- [Authorize to Microsoft Graph in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/authorize-to-microsoft-graph)
- [Register SSO AddIn in AAD v2](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2)
- [How to switch from Semi-Annual Channel to Monthly Channel for the Office 365 suite](https://docs.microsoft.com/en-us/office365/troubleshoot/administration/switch-channel-for-office-365)
- [Troubleshoot error messages for single sign-on (SSO)](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins)
- [Enable your tenant for Modern Autentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)
- [Microsoft Graph with Node/Express App](https://github.com/microsoftgraph/msgraph-training-nodeexpressapp/tree/master/Demos/03-add-msgraph)

## Copyright
Copyright (c) 2019 Microsoft. All rights reserved.
