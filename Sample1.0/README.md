# Microsoft Graph Search API Sample for .NET Core Web App

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

In this sample you will setup a .NET Core Web Application to make calls to the Microsoft Graph Search end point.  You will need access to an Azure subscription to create an Azure application and ability to add permissions to the application.

## Prerequisites

This sample requires the following:  

  * [Visual Studio 2019 or higher](https://www.visualstudio.com/en-us/downloads) 
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with sample

 1. Download or clone this repo.

### Create an Azure AD Application

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Configure the project

1.  Open the **/Sample1.0/GraphSearchApi/GraphSearchApi.sln** solution with Visual Studio
2.  Open the **appsettings.json** file
3.  Replace the values for the following:

- Domain - your Azure AD Tenant Domain
- TenantId - your Azure AD Tenant ID
- ClientId- the client id from the Configuring Azure steps
- ClientSecret - the client secret from the Configuring Azure steps

## Test the Application

1.  Press **F5**
1.  Enter a search term in the text box
1.  Select the type of data you would like to search

![Enter a search term.](./media/s01_form.png 'Perform a Search')

1.  Click **Search**, you should see the results populate on the page
1.  You can also check the **debug** checkbox to see the request/response and other debug output from the api calls

## Code of note

- The **GraphController.cs** contains the code to exchange the identity token for the graph token.  This class utilizes the **GraphAuthProvider.cs** class to do the heavy lifting.
- The **wwwroot/scripts/site.js** file contains the code that makes calls to the Graph Search API and then format it on the page

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

## Copyright
Copyright (c) 2019 Microsoft. All rights reserved.