---
title: Cell.ColumnIndex Property (Word)
keywords: vbawd10.chm156106757
f1_keywords:
- vbawd10.chm156106757
ms.prod: word
api_name:
- Word.Cell.ColumnIndex
ms.assetid: cb30b08a-b95f-da3f-ceae-7c83a5d2ec9e
ms.date: 06/08/2017
---


# Cell.ColumnIndex Property (Word)

Returns the number of the table column that contains the specified cell. Read-only  **Long** .


## Syntax

 _expression_ . **ColumnIndex**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example creates a table in a new document, selects each cell in the first row, and returns the column number that contains the selected cell.


```vb
Dim docNew As Document 
Dim tableNew As Table 
Dim cellLoop As Cell 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 
For Each cellLoop In tableNew.Rows(1).Cells 
 cellLoop.Select 
 MsgBox "This is column " &; cellLoop.ColumnIndex 
Next cellLoop
```

This example returns the column number of the cell that contains the insertion point.




```
Msgbox Selection.Cells(1).ColumnIndex
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

