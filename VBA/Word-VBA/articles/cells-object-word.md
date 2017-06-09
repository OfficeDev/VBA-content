---
title: Cells Object (Word)
ms.prod: word
ms.assetid: ceaa5b45-518d-d6ea-1ce8-5a34f6e37046
ms.date: 06/08/2017
---


# Cells Object (Word)

A collection of  **[Cell](cell-object-word.md)** objects in a table column, table row, selection, or range.


## Remarks

Use the  **Cells** property to return the **Cells** collection. The following example formats the cells in the first row in table one in the active document to be 30 points wide.


```
ActiveDocument.Tables(1).Rows(1).Cells.Width = 30
```

The following example returns the number of cells in the current row.




```
num = Selection.Rows(1).Cells.Count
```

Use the  **[Add](cells-add-method-word.md)** method to add a **[Cell](cell-object-word.md)** object to the **Cells** collection. You can also use the **[InsertCells](selection-insertcells-method-word.md)** method of the **[Selection](selection-object-word.md)** object to insert new cells. The following example adds a cell before the first cell in myTable.




```
Set myTable = ActiveDocument.Tables(1) 
myTable.Range.Cells.Add BeforeCell:=myTable.Cell(1, 1)
```

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (index), where index is the index number, to return a **Cell** object. The following example applies shading to the second cell in the first row in table one.




```
Set myCell = ActiveDocument.Tables(1).Cell(Row:=1, Column:=2) 
myCell.Shading.Texture = wdTexture20Percent
```

The following example applies shading to the first cell in the first row.




```
ActiveDocument.Tables(1).Rows(1).Cells(1).Shading _ 
 .Texture = wdTexture20Percent
```

Remarks

Use the  **Add** method with the **[Rows](rows-object-word.md)** or **[Columns](columns-object-word.md)** collection to add a row or column of cells. The following example adds a column to the first table in the active document and then inserts numbers into the first column.




```
Set myTable = ActiveDocument.Tables(1) 
Set aColumn = myTable.Columns.Add(BeforeColumn:=myTable.Columns(1)) 
For Each aCell In aColumn.Cells 
 aCell.Range.Delete 
 aCell.Range.InsertAfter num + 1 
 num = num + 1 
Next aCell
```


## Methods



|**Name**|
|:-----|
|[Add](cells-add-method-word.md)|
|[AutoFit](cells-autofit-method-word.md)|
|[Delete](cells-delete-method-word.md)|
|[DistributeHeight](cells-distributeheight-method-word.md)|
|[DistributeWidth](cells-distributewidth-method-word.md)|
|[Item](cells-item-method-word.md)|
|[Merge](cells-merge-method-word.md)|
|[SetHeight](cells-setheight-method-word.md)|
|[SetWidth](cells-setwidth-method-word.md)|
|[Split](cells-split-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](cells-application-property-word.md)|
|[Borders](cells-borders-property-word.md)|
|[Count](cells-count-property-word.md)|
|[Creator](cells-creator-property-word.md)|
|[Height](cells-height-property-word.md)|
|[HeightRule](cells-heightrule-property-word.md)|
|[NestingLevel](cells-nestinglevel-property-word.md)|
|[Parent](cells-parent-property-word.md)|
|[PreferredWidth](cells-preferredwidth-property-word.md)|
|[PreferredWidthType](cells-preferredwidthtype-property-word.md)|
|[Shading](cells-shading-property-word.md)|
|[VerticalAlignment](cells-verticalalignment-property-word.md)|
|[Width](cells-width-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
