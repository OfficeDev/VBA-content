---
title: CellRange.Borders Property (PowerPoint)
keywords: vbapp10.chm627004
f1_keywords:
- vbapp10.chm627004
ms.prod: powerpoint
api_name:
- PowerPoint.CellRange.Borders
ms.assetid: 06bd16b9-8d3e-d818-cdf4-44e0dfbaca5c
ms.date: 06/08/2017
---


# CellRange.Borders Property (PowerPoint)

Returns a  **[Borders](borders-object-powerpoint.md)** collection that represents the borders and diagonal lines for the specified **Cell** object or **CellRange** collection. Read-only.


## Syntax

 _expression_. **Borders**

 _expression_ A variable that represents a **CellRange** object.


### Return Value

Borders


## Example

This example sets the thickness of the left border for the first cell in the second row of the selected table to three points.


```vb
ActiveWindow.Selection.ShapeRange.Table.Rows(2) _
    .Cells(1).Borders.Item(ppBorderLeft).Weight = 3
```


## See also


#### Concepts


[CellRange Object](cellrange-object-powerpoint.md)

