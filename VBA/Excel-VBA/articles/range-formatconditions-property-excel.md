---
title: Range.FormatConditions Property (Excel)
keywords: vbaxl10.chm144226
f1_keywords:
- vbaxl10.chm144226
ms.prod: excel
api_name:
- Excel.Range.FormatConditions
ms.assetid: 676ffcc6-f08d-9f91-78af-7b98f8b77dca
ms.date: 06/08/2017
---


# Range.FormatConditions Property (Excel)

Returns a  **[FormatConditions](formatconditions-object-excel.md)** collection that represents all the conditional formats for the specified range. Read-only.


## Syntax

 _expression_ . **FormatConditions**

 _expression_ A variable that represents a **Range** object.


## Example

This example modifies an existing conditional format for cells E1:E10.


```vb
Worksheets(1).Range("e1:e10").FormatConditions(1) _ 
 .Modify xlCellValue, xlLess, "=$a$1"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

