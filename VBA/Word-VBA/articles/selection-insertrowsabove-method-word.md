---
title: Selection.InsertRowsAbove Method (Word)
keywords: vbawd10.chm158663195
f1_keywords:
- vbawd10.chm158663195
ms.prod: word
api_name:
- Word.Selection.InsertRowsAbove
ms.assetid: f5387043-34d0-cd84-6550-bfd96bf661b8
ms.date: 06/08/2017
---


# Selection.InsertRowsAbove Method (Word)

Inserts rows above the current selection.


## Syntax

 _expression_ . **InsertRowsAbove**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Microsoft Word inserts as many rows as there are in the current selection.

To use this method, the current selection must be in a table.


## Example

This example selects the second row in the first table and inserts a new row above it.


```vb
ActiveDocument.Tables(1).Rows(2).Select 
Selection.InsertRowsAbove
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

