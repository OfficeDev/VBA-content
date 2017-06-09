---
title: Cells.VerticalAlignment Property (Word)
keywords: vbawd10.chm155845712
f1_keywords:
- vbawd10.chm155845712
ms.prod: word
api_name:
- Word.Cells.VerticalAlignment
ms.assetid: c60fcbdb-b443-6b5a-8dd2-1c4c1e4a71d4
ms.date: 06/08/2017
---


# Cells.VerticalAlignment Property (Word)

Returns or sets the vertical alignment of text in one or more cells of a table. Read/write  **WdCellVerticalAlignment** .


## Syntax

 _expression_ . **VerticalAlignment**

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


## Example

This example creates a 3x3 table in a new document and assigns a sequential cell number to each cell in the table. The example then sets the height of the first row to 20 points and vertically aligns the text at the top of the cells.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 3, 3) 
i = 1 
For Each c In myTable.Range.Cells 
 c.Range.InsertAfter "Cell " &; i 
 i = i + 1 
Next 
With myTable.Rows(1) 
 .Height = 20 
 .Cells.VerticalAlignment = wdAlignVerticalTop 
End With
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

