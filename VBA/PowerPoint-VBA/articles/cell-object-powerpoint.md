---
title: Cell Object (PowerPoint)
keywords: vbapp10.chm628000
f1_keywords:
- vbapp10.chm628000
ms.prod: powerpoint
api_name:
- PowerPoint.Cell
ms.assetid: e89e5d69-33b1-d7b1-0a6c-4dfd8b676977
ms.date: 06/08/2017
---


# Cell Object (PowerPoint)

Represents a table cell. The  **Cell** object is a member of the **[CellRange](http://msdn.microsoft.com/library/f0914f0d-74f5-9c16-3744-efcf5c2cc36d%28Office.15%29.aspx)** collection. The **CellRange** collection represents all the cells in the specified column or row. To use the **CellRange** collection, use the **Cells** keyword.


## Remarks

You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the  **Add** method of the **Columns** or **Rows** collections to add a column or row to a table. Use the **Delete** method of the **Columns** or **Rows** collections to delete a column or row from a table.


## Example

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (index), where index is the number of the cell in the specified row or column, to return a single **Cell** object. Cells are numbered from left to right in rows and from top to bottom in columns. With right-to-left language settings, this scheme is reversed. The following example merges the first two cells in row one of the table in shape five on slide two.


```
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```

This example sets the bottom border for cell one in the first column of the table to a dashed line style.




```
With ActivePresentation.Slides(2).Shapes(5).Table.Columns(1) _

        .Cells(1)

    .Borders(ppBorderBottom).DashStyle = msoLineDash

End With
```

Use the [Shape](http://msdn.microsoft.com/library/942f67bd-b4ef-3f1f-153a-5a55aaa5663c%28Office.15%29.aspx)property to access the  **Shape** object and to manipulate the contents of each cell. This example deletes the text in the first cell (row 1, column 1), inserts new text, and then sets the width of the entire column to 110 points.




```
With ActivePresentation.Slides(2).Shapes(5).Table.Cell(1, 1)

    .Shape.TextFrame.TextRange.Delete

    .Shape.TextFrame.TextRange.Text = "Rooster"

    .Parent.Columns(1).Width = 110

End With
```


## Methods



|**Name**|
|:-----|
|[Merge](http://msdn.microsoft.com/library/e4830df1-4db9-f1e0-a4c6-d4ed2d99b9fa%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/8eef42ab-b3d1-5460-95bb-f14cbce9f434%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/edd81309-f0de-da70-67b2-4197059378fc%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/d91a9508-33a3-1b95-1786-2ab84a14ee43%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/1c9e2d38-237b-4c86-1135-af7533876501%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/45650dd8-b51d-68ec-d117-5ddb8e8c675f%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/3773ff08-043d-2b57-25ea-ba44cc30c77a%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/942f67bd-b4ef-3f1f-153a-5a55aaa5663c%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
