---
title: Application.DDEInitiate Method (Access)
keywords: vbaac10.chm12539
f1_keywords:
- vbaac10.chm12539
ms.prod: access
api_name:
- Access.Application.DDEInitiate
ms.assetid: 7b05c3ad-574e-d904-5d50-ff646486ef07
ms.date: 06/08/2017
---


# Application.DDEInitiate Method (Access)

You can use the  **DDEInitiate** function to begin a dynamic data exchange (DDE) conversation with another application. The **DDEInitiate** function opens a DDE channel for transfer of data between a DDE server and client application.


## Syntax

 _expression_. **DDEInitiate**( ** _Application_**, ** _Topic_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Application_|Required|**String**|A string expression identifying an application that can participate in a DDE conversation. Usually, the  _application_ argument is the name of an .exe file (without the .exe extension) for a Microsoft Windows?based application, such as Microsoft Excel.|
| _Topic_|Required|**String**|A string expression that is the name of a topic recognized by the  _application_ argument. Check the application's documentation for a list of topics.|

### Return Value

Variant


## Remarks

For example, if you wish to transfer data from a Microsoft Excel spreadsheet to a Microsoft Access database, you can use the  **DDEInitiate** function to open a channel between the two applications. In this example, Microsoft Access acts as the client application and Microsoft Excel acts as the server application.

If successful, the  **DDEInitiate** function begins a DDE conversation with the application and topic specified by the _application_ and _topic_ arguments, and then returns a **Long** integer value. This return value represents a unique channel number identifying a channel through which data transfer can take place. This channel number is subsequently used with other DDE functions and statements.

If the application isn't already running or if it's running but doesn't recognize the  _topic_ argument or doesn't support DDE, the **DDEInitiate** function returns a run-time error.

The value of the  _topic_ argument depends on the application specified by the _application_ argument. For applications that use documents or data files, valid topic names often include the names of those files.


 **Note**  The maximum number of channels that can be open simultaneously is determined by Microsoft Windows and your computer's memory and resources. If you aren't using a channel, you should conserve resources by terminating it with a  **DDETerminate** or **DDETerminateAll** statement.

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation.


## See also


#### Concepts


[Application Object](application-object-access.md)

