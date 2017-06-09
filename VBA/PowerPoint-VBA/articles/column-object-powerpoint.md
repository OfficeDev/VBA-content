---
title: Column Object (PowerPoint)
keywords: vbapp10.chm624000
f1_keywords:
- vbapp10.chm624000
ms.prod: powerpoint
api_name:
- PowerPoint.Column
ms.assetid: 4f289477-abab-a99a-21af-df3950b6654d
ms.date: 06/08/2017
---


# Column Object (PowerPoint)

Represents a table column. The  **Column** object is a member of the **[Columns](columns-object-powerpoint.md)** collection. The **Columns** collection includes all the columns in a table.


## Example

Use  **Columns** (index) to return a single **Column** object. Index represents the position of the column in the **Columns** collection (usually counting from left to right; although the **[TableDirection](table-tabledirection-property-powerpoint.md)** property can reverse this). This example selects the first column of the table in shape five on the second slide.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Select
```

Use the  **Cell** object to indirectly reference the **Column** object. This example deletes the text in the first cell (row 1, column 1), inserts new text, and then sets the width of the entire column to 110 points.




```vb
With ActivePresentation.Slides(2).Shapes(5).Table.Cell(1, 1)

    .Shape.TextFrame.TextRange.Delete

    .Shape.TextFrame.TextRange.Text = "Rooster"

    .Parent.Columns(1).Width = 110

End With
```

Use the  **[Add](columns-add-method-powerpoint.md)** method to add a column to a table. This example creates a column in an existing table and sets the column width to 72 points (one inch).




```vb
With ActivePresentation.Slides(2).Shapes(5).Table

    .Columns.Add.Width = 72

End With
```

Use the  **[Cells](column-cells-property-powerpoint.md)** property to modify the individual cells in a **Column** object. This example selects the first column in the table and applies a dashed line style to the bottom border.




```vb
ActiveWindow.Selection.ShapeRange.Table.Columns(1) _
    .Cells.Borders(ppBorderBottom).DashStyle = msoLineDash
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

