---
title: Cell.Selected Property (PowerPoint)
keywords: vbapp10.chm628008
f1_keywords:
- vbapp10.chm628008
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Selected
ms.assetid: 3773ff08-043d-2b57-25ea-ba44cc30c77a
ms.date: 06/08/2017
---


# Cell.Selected Property (PowerPoint)

Returns  **True** if the specified table cell is selected. Read-only.


## Syntax

 _expression_. **Selected**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Boolean


## Example

This example puts a border around the first cell in the specified table if the cell is selected.


```vb
Sub IsCellSelected()

    Dim celSelected As Cell

    Set celSelected = ActivePresentation.Slides(1).Shapes(1) _
        .Table.Columns(1).Cells(1)

    If celSelected.Selected Then
        With celSelected
            .Borders(ppBorderTop).DashStyle = msoLineRoundDot
            .Borders(ppBorderBottom).DashStyle = msoLineRoundDot
            .Borders(ppBorderLeft).DashStyle = msoLineRoundDot
            .Borders(ppBorderRight).DashStyle = msoLineRoundDot
        End With
    End If

End Sub
```


## See also


#### Concepts


[Cell Object](cell-object-powerpoint.md)

