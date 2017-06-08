---
title: FormatCondition.Borders Property (Excel)
keywords: vbaxl10.chm512079
f1_keywords:
- vbaxl10.chm512079
ms.prod: excel
api_name:
- Excel.FormatCondition.Borders
ms.assetid: 2f165a74-0b95-6643-5bd2-6a778523a411
ms.date: 06/08/2017
---


# FormatCondition.Borders Property (Excel)

Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).


## Syntax

 _expression_ . **Borders**

 _expression_ A variable that represents a **FormatCondition** object.


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


[FormatCondition Object](formatcondition-object-excel.md)

