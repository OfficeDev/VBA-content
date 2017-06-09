---
title: Columns.AutoFit Method (Word)
keywords: vbawd10.chm155910346
f1_keywords:
- vbawd10.chm155910346
ms.prod: word
api_name:
- Word.Columns.AutoFit
ms.assetid: e6a62b02-56e6-55af-9a01-12665fca3dcd
ms.date: 06/08/2017
---


# Columns.AutoFit Method (Word)

Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.


## Syntax

 _expression_ . **AutoFit**

 _expression_ Required. A variable that represents a **[Columns](columns-object-word.md)** collection.


## Remarks

If the table is already as wide as the distance between the left and right margins, this method has no affect.


## Example

This example creates a 3x3 table in a new document and then changes the width of the first column to accommodate the width of the text.


```vb
Dim docNew as Document 
Dim tableNew as Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
With tableNew 
 .Cell(1,1).Range.InsertAfter "First cell" 
 .Columns(1).AutoFit 
End With
```

This example creates a 3x3 table in a new document and then changes the width of all the columns to accommodate the width of the text.




```vb
Dim docNew as Document 
Dim tableNew as Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 
With tableNew 
 .Cell(1,1).Range.InsertAfter "First cell" 
 .Cell(1,2).Range.InsertAfter "This is cell (1,2)" 
 .Cell(1,3).Range.InsertAfter "(1,3)" 
 .Columns.AutoFit 
End With
```


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

