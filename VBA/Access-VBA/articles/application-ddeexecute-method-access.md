---
title: Application.DDEExecute Method (Access)
keywords: vbaac10.chm12540
f1_keywords:
- vbaac10.chm12540
ms.prod: access
api_name:
- Access.Application.DDEExecute
ms.assetid: 9828607e-a2e3-15e2-699a-12fb2dc9e897
ms.date: 06/08/2017
---


# Application.DDEExecute Method (Access)

You can use the  **DDEExecute** statement to send a command from a client application to a server application over an open dynamic data exchange (DDE) channel.


## Syntax

 _expression_. **DDEExecute**( ** _ChanNum_**, ** _Command_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChanNum_|Required|**Variant**|A channel number, the long integer returned by the  **[DDEInitiate](application-ddeinitiate-method-access.md)** function.|
| _Command_|Required|**String**|a command recognized by the server application. Check the server application's documentation for a list of these commands.|

## Remarks

For example, suppose you've opened a DDE channel in Microsoft Access to transfer text data from a Microsoft Excel spreadsheet into a Microsoft Access database. Use the  **DDEExecute** statement to send the **New** command to Microsoft Excel to specify that you wish to open a new spreadsheet. In this example, Microsoft Access acts as the client application, and Microsoft Excel acts as the server application.

The value of the  _command_ argument depends on the application and topic specified when the channel indicated by the _channum_ argument is opened. An error occurs if the _channum_ argument isn't an integer corresponding to an open channel or if the other application can't carry out the specified command.

From Visual Basic, you can use the  **DDEExecute** statement only to send commands to another application. For information on sending commands to Microsoft Access from another application, see[Use Microsoft Access as a DDE Server](http://msdn.microsoft.com/library/a3e82bf7-94b5-8eec-86bc-2d5387d66738%28Office.15%29.aspx).

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation.


## See also


#### Concepts


[Application Object](application-object-access.md)

