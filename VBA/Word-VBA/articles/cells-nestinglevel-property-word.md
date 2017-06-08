---
title: Cells.NestingLevel Property (Word)
keywords: vbawd10.chm155844710
f1_keywords:
- vbawd10.chm155844710
ms.prod: word
api_name:
- Word.Cells.NestingLevel
ms.assetid: 24da16e0-3713-3c74-71e9-03e886802e9f
ms.date: 06/08/2017
---


# Cells.NestingLevel Property (Word)

Returns the nesting level of the specified cells. Read-only  **Long** .


## Syntax

 _expression_ . **NestingLevel**

 _expression_ A variable that represents a **[Cells](cells-object-word.md)** collection.


## Remarks

The outermost table has a nesting level of 1. The nesting level of each successively nested table is one higher than the previous table.


## Example

This example creates a new document, creates a nested table with three levels, and then fills the first cell of each table with its nesting level.


```vb
Documents.Add 
ActiveDocument.Tables.Add Selection.Range, _ 
 3, 3, wdWord9TableBehavior, wdAutoFitContent 
With ActiveDocument.Tables(1).Range 
 .Copy 
 .Cells(1).Range.Text = .Cells.NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = .Cells.NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = _ 
 .Cells.NestingLevel 
 End With 
 End With 
End With
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

