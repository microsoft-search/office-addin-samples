# Microsoft Graph Search API Sample for Outlook AddIn (.NET)

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

TODO

## Prerequisites

This sample requires the following:  

  * [Visual Studio 2019 or higher](https://www.visualstudio.com/en-us/downloads) 
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with the sample

 1. Download or clone this repo.

### Create an Azure AD Application

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Configure the project

1.  Open the /GraphSearchApiOutlook/GraphSearchApiOutlook.sln solution with Visual Studio
2.  Open the **web.config** file
3.  Replace the values for the following:

- ida:TenantId - your Azure AD Tenant ID
- ida:ClientId- the client id from the Configuring Azure steps
- ida:Password - the client secret from the Configuring Azure steps

## Update the AddIn xml

1.  Open the **GraphSearchApiOutlook.xml** file
2.  Scroll to the bottom of the file, in the **WebApplicationInfo** section, ensure that the clientid matches the client id from the Configuring Azure steps
3.  Save the file

## Register the AddIn with Outlook

1.  Open Outlook
2.  In the ribbon, click **Get Add-ins**
3.  Click **My add-ins**, then scroll down to the **Custom add-ins** section
4.  Click **Add a custom add-in->Add from file**
5.  Browse to the **GraphSearchApiOutlook.xml** file, click **Open**
6.  Close the AddIn window
7.  You should now have a new ribbon item in a group called **Graph Search** and a button called **Search**
8.  Click the button, you should the task pane open with error.

## Test the AddIn

1.  Switch back to Visual Studio
2.  Right-click the **GraphSearchApiOutlookWeb** project, select **Debug->Start new instance**
3.  Switch back to Outlook, click **Retry** to refresh the application task pane.
4.  Run a search, review the results

>NOTE: You can set the **debug** variable in the MessageRead.js to show debugging information during the search api call process

## Code of note

- Open the **MessageRead.html** file, notice the usage of the **beta** endpoint for the Office Javascript APIs
- Open the **MessageRead.js** file, notice the code to check for the **IdentityAPI** set support
- Open the **GraphController.cs** file, notice the how the Token method accepts the identity token and trades it for a Graph Token

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
- [Enable your tenant for Modern Autentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)

## Copyright
Copyright (c) 2019 Microsoft. All rights reserved.
