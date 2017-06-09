---
title: Cell.Shape Property (PowerPoint)
keywords: vbapp10.chm628003
f1_keywords:
- vbapp10.chm628003
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Shape
ms.assetid: 942f67bd-b4ef-3f1f-153a-5a55aaa5663c
ms.date: 06/08/2017
---


# Cell.Shape Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents a shape in a table cell. Read-only.


## Syntax

 _expression_. **Shape**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Shape


## Example

This example creates a 3x3 table in a new presentation and inserts a four-pointed star into the first cell of the table.


```vb
With Presentations.Add

    With .Slides.Add(1, ppLayoutBlank)

        .Shapes.AddTable(3, 3).Select

        .Shapes(1).Table.Cell(1, 1).Shape.AutoShapeType = msoShape4pointStar

    End With

End With


```


## See also


#### Concepts


[Cell Object](cell-object-powerpoint.md)

