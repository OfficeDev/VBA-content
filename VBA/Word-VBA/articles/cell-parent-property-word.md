---
title: Cell.Parent Property (Word)
keywords: vbawd10.chm156107754
f1_keywords:
- vbawd10.chm156107754
ms.prod: word
api_name:
- Word.Cell.Parent
ms.assetid: ef27abde-9789-52f2-ac30-b346404939d6
ms.date: 06/08/2017
---


# Cell.Parent Property (Word)

Returns an  **Object** that represents the parent object of the specified **Cell** object.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example sets a variable to the first cell in the first table of the active document, changes the width of the cell to 36 points, and removes borders from the table.


```vb
Set objCell = ActiveDocument.Tables(1).Cell(1, 1) 
With objCell 
 .SetWidth ColumnWidth:=36, RulerStyle:=wdAdjustNone 
 .Parent.Borders.Enable = False 
End With
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

