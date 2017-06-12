---
title: Range.Previous Property (Excel)
keywords: vbaxl10.chm144180
f1_keywords:
- vbaxl10.chm144180
ms.prod: excel
api_name:
- Excel.Range.Previous
ms.assetid: 6ee986eb-9242-63f3-6885-1ad3730f106b
ms.date: 06/08/2017
---


# Range.Previous Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the next cell.


## Syntax

 _expression_ . **Previous**

 _expression_ A variable that represents a **Range** object.


## Remarks

If the object is a range, this property emulates pressing SHIFT+TAB; unlike the key combination, however, the property returns the previous cell without selecting it.

On a protected sheet, this property returns the previous unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the left of the specified cell.


## Example

This example selects the previous unlocked cell on Sheet1. If Sheet1 is unprotected, this is the cell immediately to the left of the active cell.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.Previous.Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

