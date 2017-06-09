---
title: Application.DDETerminate Method (Access)
keywords: vbaac10.chm12543
f1_keywords:
- vbaac10.chm12543
ms.prod: access
api_name:
- Access.Application.DDETerminate
ms.assetid: 97684f64-dd80-03b6-965d-42e9d0e6f264
ms.date: 06/08/2017
---


# Application.DDETerminate Method (Access)

You can use the  **DDETerminate** statement to close a specified dynamic data exchange (DDE) channel.


## Syntax

 _expression_. **DDETerminate**( ** _ChanNum_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChanNum_|Required|**Variant**|A channel number to close, refers to a channel opened by the  **[DDEInitiate](application-ddeinitiate-method-access.md)** function.|

### Return Value

Nothing


## Remarks

For example, if you've opened a DDE channel to transfer data between Microsoft Excel and Microsoft Access, you can use the  **DDETerminate** statement to close that channel once the transfer is complete.

If the  _channum_ argument isn't an integer corresponding to an open channel, a run-time error occurs.

Once a channel is closed, any subsequent DDE functions or statements performed on that channel cause a run-time error.

The  **DDETerminate** statement has no effect on active DDE link expressions in fields on forms or reports.

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation.


## See also


#### Concepts


[Application Object](application-object-access.md)

