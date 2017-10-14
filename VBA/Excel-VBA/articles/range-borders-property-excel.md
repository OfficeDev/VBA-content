---
title: Range.Borders Property (Excel)
keywords: vbaxl10.chm144089
f1_keywords:
- vbaxl10.chm144089
ms.prod: excel
api_name:
- Excel.Range.Borders
ms.assetid: 6d313fed-a8f0-94ba-e239-813685cd1d58
ms.date: 06/08/2017
---


# Range.Borders Property (Excel)

Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).


## Syntax

 _expression_ . **Borders**

 _expression_ A variable that represents a **Range** object.


## Example

This example sets the color of the bottom border of cell B2 on Sheet1 to a thin red border.


```vb
Sub SetRangeBorder() 
 
 With Worksheets("Sheet1").Range("B2").Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 3 
 End With 
 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

