---
title: IRtdServer.ConnectData Method (Excel)
keywords: vbaxl10.chm500006
f1_keywords:
- vbaxl10.chm500006
ms.prod: excel
api_name:
- Excel.IRtdServer.ConnectData
ms.assetid: 2d660ccc-fca7-c794-61f1-4e0578cc7511
ms.date: 06/08/2017
---


# IRtdServer.ConnectData Method (Excel)

Adds new topics from a real-time data server. The  **ConnectData** method is called when a file is opened that contains real-time data functions or when a user types in a new formula which contains the RTD function.


## Syntax

 _expression_ . **ConnectData**( **_TopicID_** , **_Strings()_** , **_GetNewValues_** )

 _expression_ A variable that represents an **IRtdServer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopicID_|Required| **Long**| A unique value, assigned by Microsoft Excel, which identifies the topic.|
| _Strings()_|Required| **Variant**|A single-dimensional array of strings identifying the topic.|
| _GetNewValues_|Required| **Boolean**| **True** to determine if new values are to be acquired.|

### Return Value

Variant


## See also


#### Concepts


[IRtdServer Object](irtdserver-object-excel.md)

