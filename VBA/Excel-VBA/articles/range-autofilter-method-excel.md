---
title: Range.AutoFilter Method (Excel)
keywords: vbaxl10.chm144084
f1_keywords:
- vbaxl10.chm144084
ms.prod: excel
api_name:
- Excel.Range.AutoFilter
ms.assetid: 0f773dbf-63e8-f714-d246-f803a74d366c
ms.date: 06/08/2017
---


# Range.AutoFilter Method (Excel)

Filters a list using the AutoFilter.


## Syntax

 _expression_ . **AutoFilter**( **_Field_** , **_Criteria1_** , **_Operator_** , **_Criteria2_** , **_VisibleDropDown_** )

 _expression_ An expression that returns a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Optional| **Variant**| The integer offset of the field on which you want to base the filter (from the left of the list; the leftmost field is field one).|
| _Criteria1_|Optional| **Variant**|The criteria (a string; for example, "101"). Use "=" to find blank fields, or use "<>" to find nonblank fields. If this argument is omitted, the criteria is All. If  _Operator_ is **xlTop10Items** , _Criteria1_ specifies the number of items (for example, "10").|
| _Operator_|Optional| **[XlAutoFilterOperator](xlautofilteroperator-enumeration-excel.md)**|One of the constants of XlAutoFilterOperator specifying the type of filter.|
| _Criteria2_|Optional| **Variant**|The second criteria (a string). Used with  _Criteria1_ and _Operator_ to construct compound criteria.|
| _VisibleDropDown_|Optional| **Variant**| **True** to display the AutoFilter drop-down arrow for the filtered field. **False** to hide the AutoFilter drop-down arrow for the filtered field. **True** by default.|

### Return Value

Variant


## Remarks

If you omit all the arguments, this method simply toggles the display of the AutoFilter drop-down arrows in the specified range.

Excel for Mac does not support this method. Similar methods on Selection and ListObject are supported.


## Example

This example filters a list starting in cell A1 on Sheet1 to display only the entries in which field one is equal to the string "Otis". The drop-down arrow for field one will be hidden.


```vb
Worksheets("Sheet1").Range("A1").AutoFilter _ 
 field:=1, _ 
 Criteria1:="Otis", _ 
 VisibleDropDown:=False
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

