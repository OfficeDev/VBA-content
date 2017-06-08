---
title: Range.Insert Method (Excel)
keywords: vbaxl10.chm144149
f1_keywords:
- vbaxl10.chm144149
ms.prod: excel
api_name:
- Excel.Range.Insert
ms.assetid: e612bbc8-3942-3349-f157-c0459805794a
ms.date: 06/08/2017
---


# Range.Insert Method (Excel)

Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.


## Syntax

 _expression_ . **Insert**( **_Shift_** , **_CopyOrigin_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shift_|Optional| **Variant**|Specifies which way to shift the cells. Can be one of the following  **[XlInsertShiftDirection](xlinsertshiftdirection-enumeration-excel.md)** constants: **xlShiftToRight** or **xlShiftDown** . If this argument is omitted, Microsoft Excel decides based on the shape of the range.|
| _CopyOrigin_|Optional| **Variant**|The copy origin.|

### Return Value

Variant


## See also


#### Concepts


[Range Object](range-object-excel.md)

