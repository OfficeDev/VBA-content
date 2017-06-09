---
title: AboveAverage.ModifyAppliesToRange Method (Excel)
keywords: vbaxl10.chm824085
f1_keywords:
- vbaxl10.chm824085
ms.prod: excel
api_name:
- Excel.AboveAverage.ModifyAppliesToRange
ms.assetid: b0bb1278-2a31-321b-1458-d62c30073610
ms.date: 06/08/2017
---


# AboveAverage.ModifyAppliesToRange Method (Excel)

Sets the cell range to which this formatting rule applies.


## Syntax

 _expression_ . **ModifyAppliesToRange**( **_Range_** )

 _expression_ A variable that represents an **AboveAverage** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **[Range](range-object-excel.md)**|The range to which this formatting rule will be applied.|

## Remarks

The range must be in the A1 reference style and be entirely contained within the sheet that is the parent of the  **[FormatConditions](formatconditions-object-excel.md)** collection. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). Dollar signs can also be used but they are ignored.

You can also use a local defined name in any part of the range, but the name must be in the language of the macro.


## See also


#### Concepts


[AboveAverage Object](aboveaverage-object-excel.md)

