---
title: CellRange Object (PowerPoint)
keywords: vbapp10.chm627000
f1_keywords:
- vbapp10.chm627000
ms.prod: powerpoint
api_name:
- PowerPoint.CellRange
ms.assetid: f0914f0d-74f5-9c16-3744-efcf5c2cc36d
ms.date: 06/08/2017
---


# CellRange Object (PowerPoint)

A collection of  **[Cell](cell-object-powerpoint.md)** objects in a table column or row. The **CellRange** collection represents all the cells in the specified column or row. To use the **CellRange** collection, use the **Cells** keyword.


## Remarks

Although the collection object is named  **CellRange** and is shown in the Object Browser, this keyword is not used in programming the PowerPoint object model. The keyword **Cells** is used instead.

You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the  **AddTable** method with the **Table** object to add a new table. Use the **Add** method of the **Columns** or **Rows** collections to add a column or row to a table. Use the **Delete** method of the **Columns** or **Rows** collections to delete a column or row from a table.


## Example

Use the  **Cells** property to return the **CellRange** collection. This example sets the right border for the cells in the first column of the table to a dashed line style.


```vb
With ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Cells

    .Borders(ppBorderRight).DashStyle = msoLineDash

End With
```

This example returns the number of cells in row one of the selected table.




```
num = ActiveWindow.Selection.ShapeRange.Table.Rows(1).Cells.Count
```

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (index), where index is the number of the cell in the specified row or column, to return a single **Cell** object. Cells are numbered from left to right in rows and from top to bottom in columns. With right-to-left language settings, this scheme is reversed. The example below merges the first two cells in row one of the table in shape five on slide two.




```vb
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

