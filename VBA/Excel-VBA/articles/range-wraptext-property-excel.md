---
title: Range.WrapText Property (Excel)
keywords: vbaxl10.chm144221
f1_keywords:
- vbaxl10.chm144221
ms.prod: excel
api_name:
- Excel.Range.WrapText
ms.assetid: 5e61b704-af16-7bad-5eeb-f163e3035513
ms.date: 06/08/2017
---


# Range.WrapText Property (Excel)

Returns or sets a  **Variant** value that indicates if Microsoft Excel wraps the text in the object.


## Syntax

 _expression_ . **WrapText**

 _expression_ A variable that represents a **Range** object.


## Remarks

This property returns  **True** if text is wrapped in all cells within the specified range, **False** if text is not wrapped in all cells within the specified range, or **Null** if the specified range contains some cells that wrap text and other cells that don't.

Microsoft Excel will change the row height of the range, if necessary, to accommodate the text in the range.


## Example

This example formats cell B2 on Sheet1 so that the text wraps within the cell.


```vb
Worksheets("Sheet1").Range("B2").Value = _ 
 "This text should wrap in a cell." 
Worksheets("Sheet1").Range("B2").WrapText = True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

