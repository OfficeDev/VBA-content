---
title: Cell.Column Property (Publisher)
keywords: vbapb10.chm5111815
f1_keywords:
- vbapb10.chm5111815
ms.prod: publisher
api_name:
- Publisher.Cell.Column
ms.assetid: 09e067a2-ee84-7a76-72b6-3b348238d020
ms.date: 06/08/2017
---


# Cell.Column Property (Publisher)

Returns a  **Long** that represents the table column containing the specified cell. Read-only.


## Syntax

 _expression_. **Column**

 _expression_A variable that represents a  **Cell** object.


## Example

This example adds a page to the active publication, creates a table on that new page, and diagonally splits all cells in even-numbered columns.


```vb
Sub CreateNewTable() 
 
 Dim pgeNew As Page 
 Dim shpTable As Shape 
 Dim tblNew As Table 
 Dim celTable As Cell 
 Dim rowTable As Row 
 
 'Creates a new document with a five-row by five-column table 
 Set pgeNew = ActiveDocument.Pages.Add(Count:=1, After:=1) 
 Set shpTable = pgeNew.Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=72, Width:=468, Height:=100) 
 Set tblNew = shpTable.Table 
 
 'Inserts a diagonal split into all cells in even-numbered columns 
 For Each rowTable In tblNew.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Column Mod 2 = 0 Then 
 celTable.Diagonal = pbTableCellDiagonalUp 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```


