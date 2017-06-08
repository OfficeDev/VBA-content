---
title: Style.Borders Property (Excel)
keywords: vbaxl10.chm177075
f1_keywords:
- vbaxl10.chm177075
ms.prod: excel
api_name:
- Excel.Style.Borders
ms.assetid: 7da8309e-f01f-b131-b462-f974dde67007
ms.date: 06/08/2017
---


# Style.Borders Property (Excel)

Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).


## Syntax

 _expression_ . **Borders**

 _expression_ A variable that represents a **Style** object.


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


[Style Object](style-object-excel.md)

