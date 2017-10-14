---
title: IRtdServer.ServerStart Method (Excel)
keywords: vbaxl10.chm500005
f1_keywords:
- vbaxl10.chm500005
ms.prod: excel
api_name:
- Excel.IRtdServer.ServerStart
ms.assetid: 5154105a-3618-fc8a-30b4-834f31c45023
ms.date: 06/08/2017
---


# IRtdServer.ServerStart Method (Excel)

The  **ServerStart** method is called immediately after a real-time data server is instantiated. Returns a **Long** ; negative value or zero indicates failure to start the server; positive value indicates success.


## Syntax

 _expression_ . **ServerStart**( **_CallbackObject_** )

 _expression_ A variable that represents an **IRtdServer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CallbackObject_|Required| **IRTDUpdateEvent**|The callback object.|

### Return Value

Long


## See also


#### Concepts


[IRtdServer Object](irtdserver-object-excel.md)

