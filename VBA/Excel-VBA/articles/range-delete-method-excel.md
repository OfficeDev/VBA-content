---
title: Range.Delete Method (Excel)
keywords: vbaxl10.chm144115
f1_keywords:
- vbaxl10.chm144115
ms.prod: excel
api_name:
- Excel.Range.Delete
ms.assetid: 7d890cc5-5b5b-35f9-2d97-e4fe48f244ee
ms.date: 06/08/2017
---


# Range.Delete Method (Excel)

Deletes the object.


## Syntax

 _expression_ . **Delete**( **_Shift_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shift_|Optional| **Variant**|Used only with  **[Range](range-object-excel.md)** objects. Specifies how to shift cells to replace deleted cells. Can be one of the following **[XlDeleteShiftDirection](xldeleteshiftdirection-enumeration-excel.md)** constants: **xlShiftToLeft** or **xlShiftUp** . If this argument is omitted, Microsoft Excel decides based on the shape of the range.|

### Return Value

Variant


## See also


#### Concepts


[Range Object](range-object-excel.md)

