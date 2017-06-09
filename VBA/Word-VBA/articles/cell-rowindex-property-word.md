---
title: Cell.RowIndex Property (Word)
keywords: vbawd10.chm156106756
f1_keywords:
- vbawd10.chm156106756
ms.prod: word
api_name:
- Word.Cell.RowIndex
ms.assetid: 745fabed-ba99-2e69-0d87-a7b520ac78cf
ms.date: 06/08/2017
---


# Cell.RowIndex Property (Word)

Returns the number of the row that contains the specified cell. Read-only  **Long** .


## Syntax

 _expression_ . **RowIndex**

 _expression_ An expression that returns a **[Cell](cell-object-word.md)** object.


## Example

This example creates a 3x3 table in a new document, selects each cell in the first column, and displays the row number that contains each selected cell.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
For Each aCell In myTable.Columns(1).Cells 
 aCell.Select 
 MsgBox "This is row " &; aCell.RowIndex 
Next aCell
```

This example displays the row number of the first row in the selection.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Msgbox Selection.Cells(1).RowIndex 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

