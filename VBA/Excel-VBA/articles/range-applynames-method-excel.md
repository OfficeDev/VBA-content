---
title: Range.ApplyNames Method (Excel)
keywords: vbaxl10.chm144079
f1_keywords:
- vbaxl10.chm144079
ms.prod: excel
api_name:
- Excel.Range.ApplyNames
ms.assetid: 3798ecfb-c839-64a9-1088-d7752a3e81ae
ms.date: 06/08/2017
---


# Range.ApplyNames Method (Excel)

Applies names to the cells in the specified range.


## Syntax

 _expression_ . **ApplyNames**( **_Names_** , **_IgnoreRelativeAbsolute_** , **_UseRowColumnNames_** , **_OmitColumn_** , **_OmitRow_** , **_Order_** , **_AppendLast_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Names_|Optional| **Variant**| An array of the names to be applied. If this argument is omitted, all names on the sheet are applied to the range.|
| _IgnoreRelativeAbsolute_|Optional| **Variant**| **True** to replace references with names, regardless of the reference types of either the names or references. **False** to replace absolute references only with absolute names, relative references only with relative names, and mixed references only with mixed names. The default value is **True** .|
| _UseRowColumnNames_|Optional| **Variant**| **True** to use the names of row and column ranges that contain the specified range if names for the range cannot be found. **False** to ignore the _OmitColumn_ and _OmitRow_ arguments. The default value is **True** .|
| _OmitColumn_|Optional| **Variant**| **True** to replace the entire reference with the row-oriented name. The column-oriented name can be omitted only if the referenced cell is in the same column as the formula and is within a row-oriented named range. The default value is **True** .|
| _OmitRow_|Optional| **Variant**| **True** to replace the entire reference with the column-oriented name. The row-oriented name can be omitted only if the referenced cell is in the same row as the formula and is within a column-oriented named range. The default value is **True** .|
| _Order_|Optional| **[XlApplyNamesOrder](xlapplynamesorder-enumeration-excel.md)**|Determines which range name is listed first when a cell reference is replaced by a row-oriented and column-oriented range name.|
| _AppendLast_|Optional| **Variant**| **True** to replace the definitions of the names in _Names_ and also replace the definitions of the last names that were defined. **False** to replace the definitions of the names in _Names_ only. The default value is **False** .|

### Return Value

Variant


## Remarks

You can use the  **Array** function to create the list of names for the _Names_ argument.

If you want to apply names to the entire sheet, use  `Cells.ApplyNames`.

You cannot "unapply" names; to delete names, use the  **Delete** method.


## Example

This example applies names to the entire sheet.


```
Cells.ApplyNames Names:=Array("Sales", "Profits")
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

