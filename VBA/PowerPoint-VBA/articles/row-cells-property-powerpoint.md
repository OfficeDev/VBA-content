---
title: Row.Cells Property (PowerPoint)
keywords: vbapp10.chm626003
f1_keywords:
- vbapp10.chm626003
ms.prod: powerpoint
api_name:
- PowerPoint.Row.Cells
ms.assetid: 2cbd413f-21ab-fdb1-9a88-b64af753ae4c
ms.date: 06/08/2017
---


# Row.Cells Property (PowerPoint)

Returns a  **[CellRange](cellrange-object-powerpoint.md)** collection that represents the cells in a table column or row. Read-only.


## Syntax

 _expression_. **Cells**

 _expression_ A variable that represents a **Row** object.


### Return Value

CellRange


## Example

This example creates a new presentation, adds a slide, inserts a 3x3 table on the slide, and assigns the column and row number to each cell in the table.


```vb
Dim i As Integer

Dim j As Integer

With Presentations.Add

    .Slides.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select
    Set myTable = .Slides(1).Shapes(1).Table

    For i = 1 To myTable.Columns.Count
        For j = 1 To myTable.Columns(i).Cells.Count
            myTable.Columns(i).Cells(j).Shape.TextFrame _
                .TextRange.Text = "col. " &; i &; "row " &; j
        Next j
    Next i

End With
```


## See also


#### Concepts


[Row Object](row-object-powerpoint.md)

