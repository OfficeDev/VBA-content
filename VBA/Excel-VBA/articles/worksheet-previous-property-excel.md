---
title: Worksheet.Previous Property (Excel)
keywords: vbaxl10.chm174086
f1_keywords:
- vbaxl10.chm174086
ms.prod: excel
api_name:
- Excel.Worksheet.Previous
ms.assetid: 8409e3c6-564e-2ba1-1e49-79a1c37cc845
ms.date: 06/08/2017
---


# Worksheet.Previous Property (Excel)

Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.


## Syntax

 _expression_ . **Previous**

 _expression_ A variable that represents a **Worksheet** object.


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


[Worksheet Object](worksheet-object-excel.md)

