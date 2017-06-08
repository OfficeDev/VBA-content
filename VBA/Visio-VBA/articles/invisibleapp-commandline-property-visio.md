---
title: InvisibleApp.CommandLine Property (Visio)
keywords: vis_sdr.chm17513280
f1_keywords:
- vis_sdr.chm17513280
ms.prod: visio
api_name:
- Visio.InvisibleApp.CommandLine
ms.assetid: fb3646b4-5191-71b2-1d6c-23764e764865
ms.date: 06/08/2017
---


# InvisibleApp.CommandLine Property (Visio)

Determines how Microsoft Visio was started. Read-only.


## Syntax

 _expression_ . **CommandLine**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

String


## Remarks

When you double-click a drawing, template, or stencil icon to start the application, the name of the file appears in the string returned by the  **CommandLine** property. When you use a **CreateObject** call to start the application, "/Automation" appears in the string. When you double-click a Visio embedded object in an OLE container application, "/Embedding" appears in the string.

The following table includes other command line switches you can use to start the application.



|**Command line switch**|**Description **|
|:-----|:-----|
|/nonew|The  **New** tab is not shown on startup.|
|/nologo|The startup screen is not shown on startup.|
|/p filename|The  **Print** dialog box is shown, so that you can quickly print a file.|
|filename|Opens a Visio file. Either the file has to be in the  **Drawings** file path in the **File Locations** dialog box (click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click ** File Locations**), or you must name an absolute path.|
|/1, /2, /3,.../9|Opens one of the last-opened files.|
|/noreg|Prevents Visio from registering itself.|
|/u|Unregisters Visio.|
|/r|Registers Visio.|
|/s|Silently registers Visio.|
|/pt filename,[printername, drivername, portname]|Directs the file to print on a particular printer. (Added in Visio version 5.0c.)|
|::ODMA|Visio uses ODMA to open a file.|

