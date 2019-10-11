# Microsoft Graph Search API Sample for Excel AddIn (.NET)

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

This sample will show you how to implement an Excel AddIn that makes a call to the Microsoft Graph Search API endpoint and then display the results in a workbook.

## Prerequisites

This sample requires the following:  

  * [Visual Studio 2019 or higher](https://www.visualstudio.com/en-us/downloads) 
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with the sample

 1. Download or clone this repo.

## Configure Azure

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Configure the project

1.  Open the **Sample2.1/ExcelWebAddIn1/GraphSearchApiExcel.sln** solution with Visual Studio
2.  In the **GraphSearchApiExcelWeb** project, open the **web.config** file
3.  Replace the values for the following:

- ida:TenantId - your Azure AD Tenant ID
- ida:ClientId- the client id from the Configuring Azure steps
- ida:Password - the client secret from the Configuring Azure steps

## Update the AddIn xml

1.  Open the **GraphSearchApiExcel.xml** file
2.  Scroll to the bottom of the file, in the **WebApplicationInfo** section, ensure that the clientid matches the client id from the configuration steps above.

![Enter your application id and the uri](./media/s03_WebAppInfo.png 'Update the AddIn xml')

3.  Save the file

## Create a File Share

1.  On your development machine, create a folder called **c:\manifests**
2.  Right-click the folder, select **Properties**
3.  Click the **Sharing** tab
4.  Click **Share**

![Create a Share](./media/s03_Share.png 'Create a Share')

5.  Enter your account name, click **Share**
6.  Close the dialog
7.  Copy the **GraphSearchApiExcel.xml** file you edited above to the new manifests share

## Register the AddIn Fileshare with Office

1.  Open Excel
1.  Select **Blank Workbook**

![Create a blank workbook](./media/s03_Excel_Blank.png 'Create a blank workbook')

2.  Click **File->Options**
3.  Select the **Trust Center** tab
4.  Click the **Trust Center Settings** button

![Open the trust center](./media/s03_Excel_TrustCenter.png 'Open the trust center')

5.  Select the **Trusted Add-in Catalogs** tab
6.  In the catalog url, type **//localhost/manifests**, then click **Add catalog**

![Add a trusted catalog location](./media/s03_Excel_AddCatalog.png 'Add a trusted catalog location')

7.  Click **OK** and close out all dialogs

## Register the AddIn with Excel

1.  In the ribbon, select the **Insert** tab
2.  Click **Get Add-ins**

![Add an AddIn](./media/s03_Excel_AddAddIn.png 'Add an AddIn')

3.  Select **SHARED FOLDER**, the select the **Microsoft Graph Search** Add-in
4.  Click **Add**

![Add the Graph AddIn](./media/s03_Excel_AddGraphAddIn.png 'Add the Graph AddIn')

6.  Close the AddIn window
7.  Click the **Home** tab
8.  You should now have a new ribbon item in a group called **Search** and a button called **Graph Search API**
8.  Click the button, you should the task pane open with error:

![Explore the AddIn](./media/s03_Excel_ExploreAddIn.png 'Explore the AddIn')

## Test the AddIn

1.  Switch back to Visual Studio
2.  Right-click the **GraphSearchApiExcelWeb** project, select **Debug->Start new instance**
3.  Switch back to Excel, click **Retry** to refresh the application task pane.
4.  Run a search
5.  Review the results that are exported to the workbook sheet as a filterable and searchable table

## Code of note

- The **GraphController.cs** file is responsible for trading the identity token for a graph token.
-  The **Home.js** file contains a method called **parseResult** that will dynamically create an Excel table based on the columns returned from the search result.

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
