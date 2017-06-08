---
title: Range.Hidden Property (Excel)
keywords: vbaxl10.chm144145
f1_keywords:
- vbaxl10.chm144145
ms.prod: excel
api_name:
- Excel.Range.Hidden
ms.assetid: 7e785c38-a8ae-3810-a88a-0bfb7b74e2d6
ms.date: 06/08/2017
---


# Range.Hidden Property (Excel)

Returns or sets a  **Variant** value that indicates if the rows or columns are hidden.


## Syntax

 _expression_ . **Hidden**

 _expression_ A variable that represents a **Range** object.


## Remarks

Set this property to  **True** to hide a row or column. The specified range must span an entire column or row.

Don't confuse this property with the  **[FormulaHidden](range-formulahidden-property-excel.md)** property.


## Example

This example hides column C on Sheet1.


```vb
Worksheets("Sheet1").Columns("C").Hidden = True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

