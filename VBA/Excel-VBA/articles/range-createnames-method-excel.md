---
title: Range.CreateNames Method (Excel)
keywords: vbaxl10.chm144108
f1_keywords:
- vbaxl10.chm144108
ms.prod: excel
api_name:
- Excel.Range.CreateNames
ms.assetid: 00c7c74f-606d-7eee-ac52-f6b21446f5be
ms.date: 06/08/2017
---


# Range.CreateNames Method (Excel)

Creates names in the specified range, based on text labels in the sheet.


## Syntax

 _expression_ . **CreateNames**( **_Top_** , **_Left_** , **_Bottom_** , **_Right_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Top_|Optional| **Variant**| **True** to create names by using labels in the top row. The default value is **False** .|
| _Left_|Optional| **Variant**| **True** to create names by using labels in the left column. The default value is **False** .|
| _Bottom_|Optional| **Variant**| **True** to create names by using labels in the bottom row. The default value is **False** .|
| _Right_|Optional| **Variant**| **True** to create names by using labels in the right column. The default value is **False** .|

### Return Value

Variant


## Remarks

If you don?t specify one of  _Top_,  _Left_,  _Bottom_, or  _Right_, Microsoft Excel guesses the location of the text labels, based on the shape of the specified range.


## Example

This example creates names for cells B1:B3 based on the text in cells A1:A3. Note that you must include the cells that contain the names in the range, even though the names are created only for cells B1:B3.


```vb
Set rangeToName = Worksheets("Sheet1").Range("A1:B3") 
rangeToName.CreateNames Left:=True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

