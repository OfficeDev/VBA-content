---
title: Cell.NestingLevel Property (Word)
keywords: vbawd10.chm156106859
f1_keywords:
- vbawd10.chm156106859
ms.prod: word
api_name:
- Word.Cell.NestingLevel
ms.assetid: 6eff7eac-72b9-1b33-af2c-0dd410576e92
ms.date: 06/08/2017
---


# Cell.NestingLevel Property (Word)

Returns the nesting level of the specified cell. Read-only  **Long** .


## Syntax

 _expression_ . **NestingLevel**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


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
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = _ 
 .Cells(1).NestingLevel 
 End With 
 End With 
End With
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

