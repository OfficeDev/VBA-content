---
title: Application.DDERequest Method (Access)
keywords: vbaac10.chm12542
f1_keywords:
- vbaac10.chm12542
ms.prod: access
api_name:
- Access.Application.DDERequest
ms.assetid: c6f5f472-aeac-6de9-8133-bebfc5887eee
ms.date: 06/08/2017
---


# Application.DDERequest Method (Access)

You can use the  **DDERequest** function over an open dynamic data exchange (DDE) channel to request an item of information from a DDE server application.


## Syntax

 _expression_. **DDERequest**( ** _ChanNum_**, ** _Item_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChanNum_|Required|**Variant**|A channel number, the integer returned by the  **DDEInitiate** function.|
| _Item_|Required|**String**|A string expression that's the name of a data item recognized by the application specified by the  **DDEInitiate** function. Check the application's documentation for a list of possible items.|

### Return Value

String


## Remarks

For example, if you have an open DDE channel between Microsoft Access and Microsoft Excel, you can use the  **DDERequest** function to transfer text from a Microsoft Excel spreadsheet to a Microsoft Access database.

The  _channum_ argument specifies the channel number of the desired DDE conversation, and the _item_ argument identifies which data should be retrieved from the server application. The value of the _item_ argument depends on the application and topic specified when the channel indicated by the _channum_ argument is opened. For example, the _item_ argument may be a range of cells in a Microsoft Excel spreadsheet.

The  **DDERequest** function returns a **Variant** as a string containing the requested information if the request was successful.

The data is requested in alphanumeric text format. Graphics or text in any other format can't be transferred.

If the  _channum_ argument isn't an integer corresponding to an open channel, or if the data requested can't be transferred, a run-time error occurs.

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation .


## See also


#### Concepts


[Application Object](application-object-access.md)

