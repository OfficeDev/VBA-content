---
title: IRtdServer.RefreshData Method (Excel)
keywords: vbaxl10.chm500007
f1_keywords:
- vbaxl10.chm500007
ms.prod: excel
api_name:
- Excel.IRtdServer.RefreshData
ms.assetid: 42a2ad6f-a413-6b09-ca38-3369475e1cd5
ms.date: 06/08/2017
---


# IRtdServer.RefreshData Method (Excel)

This method is called by Microsoft Excel to get new data. Returns a  **Variant** .


## Syntax

 _expression_ . **RefreshData**( **_TopicCount_** )

 _expression_ A variable that represents an **IRtdServer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopicCount_|Required| **Long**|The RTD server must change the value of the  **TopicCount** to the number of elements in the array returned.|

### Return Value

A Variant array that contains the new data.


## Remarks

The data returned to Excel is a  **Variant** containing a two-dimensional array. The first dimension represents the list of topic IDs. The second dimension represents the values associated with the topic IDs.


## See also


#### Concepts


[IRtdServer Object](irtdserver-object-excel.md)

