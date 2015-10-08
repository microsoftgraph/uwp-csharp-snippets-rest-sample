# Office 365 Snippets Sample for UWP Using Unified API (Preview)

**Table of contents**

* [Introduction](#introduction)
* [Prerequisites](#prerequisites)
* [Find the system-assigned redirect URI](#redirect)
* [Register and configure the app](#register)
* [Build and debug](#build)
* [How the sample affects your tenant data](#how-the-sample-affects-your-tenant-data)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="introduction"></a>
##Introduction

This sample shows how to use the unified API to send email, manage groups, and perform other activities with Office 365 data.
The Office 365 unified API exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to the Office 365 unified API in a Windows 10 universal app. 


**Note: If possible, please use this sample with a "non-work" or test account in Office 365. With the current version of the project, it does not always clean up the created objects in your mailbox and calendar. At this time you'll have to manually remove sample mails and calendar events.**  


<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  
  * Visual Studio 2015  
  * Windows 10 Tools for Visual Studio
  * Windows 10 (development mode enabled)
  * An Office 365 for business account. You can sign up for [an Office 365 Developer subscription](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_Office365Account) that includes the resources that you need to start building Office 365 apps.
  * A Microsoft Azure tenant to register your application. Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

     > Important: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see [Associate your Office 365 account with Azure AD to create and manage apps](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) for more information.
      

<a name="redirect"></a>
## Find the system-assigned redirect URI for the app

Before you can register the application in the Azure portal, you need to find out the application's redirect URI.  Windows 10 provides each application with a unique URI and ensures that messages sent to that URI are only sent to that application.  To determine the redirect URI for your project:

1. Open the solution in Visual Studio 2015. 
2. Make sure that your Platform Target is set to x86 or x64.
3. Press F5.
4. After the app launches, choose the **Copy** button ![alt text](/readme-images/copy_icon.png) located in the menu on the top left of the app. This will copy the redirect URI for the app to the clipboard. 
5. Store this value. You will use it when registering the app, as described in the following section. 


The redirect URI value will look something like this:
```
ms-appx-web://Microsoft.AAD.BrokerPlugIn/S-1-15-2-694665007-945573255-503870805-3898041910-4166806349-50292026-2305040851
```


<a name="register"></a>
##Register and configure the app

1.	Sign in to the [Azure Management Portal](http://aka.ms/i5b8dz) using your Azure AD credentials.
2.	Click **Active Directory** on the left menu, then select the directory for your Office 365 developer site.
3.	On the top menu, click **Applications**.
4.	Click **Add** from the bottom menu.
5.	On the **What do you want to do page**, click **Add an application my organization is developing**.
6.	On the **Tell us about your application page**, select **NATIVE CLIENT APPLICATION** for type and specify a name for the app, for example **O365-UWP-Snippets**.
7.	Click the arrow icon on the lower-right corner of the page.
8.	On the **Application information** page, enter the redirect URI value that you obtained during the previous step.
9.	Once the application is successfully added, you'll be taken to the **Quick Start** page for the application. From there, select **Configure** in the top menu.
10.	Under **permissions to other applications**, select **Add application**. In the dialog box, select the **Office 365 unified API (preview)** application. 
11.	Select the following permissions: 
	* Read and write signed-in user's calendars
	* Read signed-in user's contacts
	* Read signed-in user's files
	* Send mail as signed-in user
	* Read signed-in user's mail
	* Read all users' full profiles
	* Read and write signed-in user's profile
	* Access directory as the signed-in user
	* Read and write directory data
12.	Copy the value specified for **Client ID** on the **Configure** page.
13.	Click **Save** in the bottom menu.

**Note**: The **Read and write directory data** permission allows the app to create a user in the tenant. Do not set this permission unless you intend to run the sample with an admin account. The **Create User** snippet will return a 403 (Forbidden) status code if you run the sample with a non-admin account.

<a name="build"></a>
## Build and debug ##

**Note:** If you see any errors while installing packages during step 2, make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue.

1. After you've loaded the solution in Visual Studio, configure the sample to use the client id that you registered in Azure Active Directory and the domain of your tenant by adding the corresponding values for these keys in the Application.Resources node of the App.xaml file.
![Office 365 UWP unified API snippets sample](/readme-images/ClientTenant.png "Client ID value in App.xaml file")`

2. Press F5 to build and debug. Run the solution and sign in to Office 365 with your organizational account.

<a name="#how-the-sample-affects-your-tenant-data"></a>
##How the sample affects your tenant data
This sample runs REST commands that create, read, update, or delete data. When running commands that delete or edit data, the sample creates fake entities. The fake entities are deleted or edited so that your actual tenant data is unaffected. The sample will leave behind fake entities on your tenant.

<a name="questions"></a>
## Questions and comments

We'd love to get your feedback about the O365 UWP unified API Snippets project. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/O365-UWP-Unified-API-Snippets/issues) section of this repository.

Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [API].

<a name="additional-resources"></a>
## Additional resources ##

- [Other Office 365 Connect samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Connect)
- [Office 365 unified API overview (preview)](https://msdn.microsoft.com/en-us/office/office365/howto/office-365-unified-api-overview)
- [Office 365 APIs platform overview](https://msdn.microsoft.com/office/office365/howto/platform-development-overview)
- [Office 365 API code samples and videos](https://msdn.microsoft.com/office/office365/howto/starter-projects-and-code-samples)
- [Office developer code samples](http://dev.office.com/code-samples)
- [Office dev center](http://dev.office.com/)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.


