---
title: Range.Cut Method (Excel)
keywords: vbaxl10.chm144112
f1_keywords:
- vbaxl10.chm144112
ms.prod: excel
api_name:
- Excel.Range.Cut
ms.assetid: b9f525c4-c314-450c-f88b-e6c5cdc00112
ms.date: 06/08/2017
---


# Range.Cut Method (Excel)

Cuts the object to the Clipboard or pastes it into a specified destination.


## Syntax

 _expression_ . **Cut**( **_Destination_** )

 _expression_ An expression that returns a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Destination_|Optional| **Variant**|The range where the object should be pasted. If this argument is omitted, the object is cut to the Clipboard.|

### Return Value

Variant


## Remarks

The cut range must be made up of adjacent cells.


## See also


#### Concepts


[Range Object](range-object-excel.md)

