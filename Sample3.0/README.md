# Microsoft Graph Search API Sample for SPFx

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

This sample demonstrates how to make calls to the Microsoft Graph Search API using SPFx web parts.

## Prerequisites

This sample requires the following:  

  * [Visual Studio Code](TODO) 
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with the sample

 1. Download or clone this repo.
 2. Open Visual Studio code to the Sample Repo directory
 3. Run the following commands to setup your development environment

 ```Javascript
npm install -g yo gulp
npm install -g @microsoft/generator-sharepoint
gulp trust-dev-cert
 ```
 
### Create an Azure AD Application

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Configure the project

1.  Open Visual Studio code to the ./Sample3.0/spfx folder

## Test the Web Part (Local)

1.  In the debug console, type the following:

```javascript
cd sample-3
gulp serve
```

2. A window will open to the local workbench, click the **+** sign.

3.  Select the Graph Search API web part

4.  Click the edit icon, then select **MSGraphClient**

5.  Type a search term, then click **Search**

## Deploy and Assign Permissions (SharePoint Online)

1.  In the debug console, type the following:

```javascript
gulp package-solution
```

2. Open a new browser window to your (SharePoint Online Administration site)[https://YOURTENANT-admin.sharepoint.com]

3.  Click **Active Sites**, ensure that you have an app catalog template site created

>NOTE:  If you do not have one, you will need to create one and wait for 20-30 minutes for it to completely provision.  Failure to do so will require you to deploy your web part several times.

4.  Open the App Catalog site, then click **Apps for SharePoint**

5.  Click **Upload**, browse to the ./sharepoint.solution directory and select the **sample-3.sppkg** file

6.  Switch back to the SharePoint Online admin center, click **API Management**

>NOTE:  API Management will not display until your App Catalog has been created and the backend has converged.

7.  Approve all the permissions that your application has requested.

##  Add web part to a page

1.  Open a different SharePoint site, click the **Pages** document library

2.  Click **+New->Web Part Page** in the menu

3.  For the name, type "GraphSearch", then click **Create**

4.  In the **Header** web part zone, click **Add a Web Part**

5. Select the **Other** category, and then select the **GraphSearchAPI** web part, then click **Add**

6.  In the ribbon, click **Edit page**

7.  In the web part drop down, select **Edit web part**, then click **Configure**

8.  Select **MSGraphClient** and then close the property window

9.  Click **OK** to save the web part properties

10.  Type a search term, then click **Search**, you will see the Graph Search API results display in the area below the web part.

## Code of note

- The **package-solution.json** file contains the permissions that are needed for your web part.

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
- [Consume the Microsoft Graph in the SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial)

## Copyright
Copyright (c) 2019 Microsoft. All rights reserved.
