---
title: Selection.InsertRowsBelow Method (Word)
keywords: vbawd10.chm158663193
f1_keywords:
- vbawd10.chm158663193
ms.prod: word
api_name:
- Word.Selection.InsertRowsBelow
ms.assetid: d36441d1-ff1f-b557-d0d0-1d12d4abab2d
ms.date: 06/08/2017
---


# Selection.InsertRowsBelow Method (Word)

Inserts rows below the current selection.


## Syntax

 _expression_ . **InsertRowsBelow**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Microsoft Word inserts as many rows as there are in the current selection.

To use this method, the current selection must be in a table.


## Example

This example selects the second row in the first table and inserts a new row below it.


```vb
ActiveDocument.Tables(1).Rows(2).Select 
Selection.InsertRowsBelow
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

