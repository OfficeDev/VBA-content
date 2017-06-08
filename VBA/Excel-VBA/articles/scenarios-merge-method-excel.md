---
title: Scenarios.Merge Method (Excel)
keywords: vbaxl10.chm362077
f1_keywords:
- vbaxl10.chm362077
ms.prod: excel
api_name:
- Excel.Scenarios.Merge
ms.assetid: db956914-aec1-ed2a-e4fa-d0f9c15ec882
ms.date: 06/08/2017
---


# Scenarios.Merge Method (Excel)

Merges the scenarios from another sheet into the  **[Scenarios](scenarios-object-excel.md)** collection.


## Syntax

 _expression_ . **Merge**( **_Source_** )

 _expression_ A variable that represents a **Scenarios** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The name of the sheet that contains scenarios to be merged, or a  **[Worksheet](worksheet-object-excel.md)** object that represents that sheet.|

### Return Value

Variant


## Remarks

The value of a merged range is specified in the cell of the range's upper-left corner.


## See also


#### Concepts


[Scenarios Object](scenarios-object-excel.md)

