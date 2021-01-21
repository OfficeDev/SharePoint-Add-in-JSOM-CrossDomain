---
page_type: sample
products:
- office-sp
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - OAuth 2.0
  createdDate: 8/12/2015 1:15:58 PM
description: "Use the SharePoint Cross Domain JavaScript library to access SharePoint data from a remotely hosted web page without the need for OAuth tokens."
---

# Access SharePoint data with the Cross Domain JavaScript Library

> SharePoint add-in model is considered as a legacy option for extending SharePoint user interface. Please see [SharePoint Framework documentation](https://aka.ms/spfx) and the [SharePoint Framework samples](https://aka.ms/spfx-webparts) for the future proven option to extend SharePoint Online. Possible backend services should be using Azure Active Directly based registration and related app models.

## Summary

Use the SharePoint Cross Domain JavaScript library (CDL) to access SharePoint data from a remotely hosted web page without the need for OAuth tokens.

### Applies to

-  SharePoint Online and on-premise SharePoint 2013 and later 

## Prerequisites

This sample requires the following:

- A SharePoint 2013 (or later) development environment that is configured for add-in isolation. (A SharePoint Online Developer Site is automatically configured. For an on premise development environment, see [Set up an on-premises development environment for SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179923.aspx)) 

- Visual Studio and the Office Developer Tools for Visual Studio installed on your developer computer 

## Description of the code

The sample includes an instance of an Announcements list with two sample announcements in it. The list instance is deployed to the add-in web. When the start page loads, JavaScript in the CrossDomainExec.js file, in the remote domain, makes an AJAX call to SharePoint to get the announcements and displays them on the page. 

The page looks similar to the following.

![The add-in start page with a table showing the result of searching on the term "SharePoint".](/description/image.png) 

**Note:** The start page of the add-in, CrossDomainCall.aspx, is a plain HTML file. It has an "aspx" extension because when a remote start page in a provider-hosted add-in is opened by SharePoint, SharePoint sends it a context token in the body of the HTTP request. To do this SharePoint must send a POST request. But the IIS Express that is hosting the remote page when you are debugging will not accept POST requests to a resource with a .html extension. 

The sample demonstrates the following:

- How to use the key classes of the CDL to make calls from a remotely hosted domain to the SharePoint add-in web domain. 

- How to parse the JSON-formatted data returned from the SharePoint and how to display it dynamically. 

## To use the sample

12. Open **Visual Studio** as an administrator.
13. Open the .sln file.
13. In **Solution Explorer**, highlight the SharePoint add-in project and replace the **Site URL** property with the URL of your SharePoint developer site.
14. Press F5.
15. After the add-in installs, the consent page opens. Click **Trust It**.
16. The page opens that the JavaScript runs immediately.


## Troubleshooting

For troubleshooting steps, visit the "Troubleshooting the solution" table in the [cross-domain library documentation article](http://msdn.microsoft.com/library/bc37ff5c-1285-40af-98ae-01286696242d).

## Questions and comments

We'd love to get your feedback on this sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/SharePoint-Add-in-JSOM-CrossDomain/issues) section of this repository.

<a name="resources"/>
## Additional resources

- [SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179930.aspx )
- [Creating SharePoint Add-ins that use the cross-domain library](https://msdn.microsoft.com/library/office/dn790708.aspx)
- [Access SharePoint 2013 data from add-ins using the cross-domain library](https://msdn.microsoft.com/library/office/fp179927.aspx)

### Copyright

Copyright (c) Microsoft. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
