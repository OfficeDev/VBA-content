---
title: Cell Object (Word)
keywords: vbawd10.chm2382
f1_keywords:
- vbawd10.chm2382
ms.prod: word
api_name:
- Word.Cell
ms.assetid: cbe6ae71-b2da-63a9-1446-0a2f81ab8b14
ms.date: 06/08/2017
---


# Cell Object (Word)

Represents a single table cell. The  **Cell** object is a member of the **[Cells](cells-object-word.md)** collection. The **Cells** collection represents all the cells in the specified object.


## Remarks

Use  **[Cell](table-cell-method-word.md)** (row, column), where row is the row number and column is the column number, or **Cells** (index), where index is the index number, to return a **Cell** object. The following example applies shading to the second cell in the first row.


```vb
Set myCell = ActiveDocument.Tables(1).Cell(Row:=1, Column:=2) 
myCell.Shading.Texture = wdTexture20Percent
```

The following example applies shading to the first cell in the first row.




```vb
ActiveDocument.Tables(1).Rows(1).Cells(1).Shading _ 
 .Texture = wdTexture20Percent
```

Use the  **[Add](cells-add-method-word.md)** method to add a **Cell** object to the **[Cells](cells-object-word.md)** collection. You can also use the **[InsertCells](selection-insertcells-method-word.md)** method of the **[Selection](selection-object-word.md)** object to insert new cells. The following example adds a cell before the first cell in `myTable`.




```vb
Set myTable = ActiveDocument.Tables(1) 
myTable.Range.Cells.Add BeforeCell:=myTable.Cell(1, 1)
```

The following example sets a range ( _myRange_ ) that references the first two cells in the first table. After the range is set, the cells are combined by the **[Merge](cells-merge-method-word.md)** method.




```vb
Set myTable = ActiveDocument.Tables(1) 
Set myRange = ActiveDocument.Range(myTable.Cell(1, 1) _ 
 .Range.Start, myTable.Cell(1, 2).Range.End) 
myRange.Cells.Merge
```

Remarks

Use the  **[Add](addins-add-method-word.md)** method with the **[Rows](rows-object-word.md)** or **[Columns](columns-object-word.md)** collection to add a row or column of cells.

Use the  **[Information](selection-information-property-word.md)** property with a **Selection** object to return the current row and column number. The following example changes the width of the first cell in the selection and then displays the cell's row number and column number.




```vb
If Selection.Information(wdWithInTable) = True Then 
 With Selection 
 .Cells(1).Width = 22 
 MsgBox "Cell " &; .Information(wdStartOfRangeRowNumber) _ 
 &; "," &; .Information(wdStartOfRangeColumnNumber) 
 End With 
End If
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


