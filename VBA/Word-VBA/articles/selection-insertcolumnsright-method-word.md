---
title: Selection.InsertColumnsRight Method (Word)
keywords: vbawd10.chm158663194
f1_keywords:
- vbawd10.chm158663194
ms.prod: word
api_name:
- Word.Selection.InsertColumnsRight
ms.assetid: 0367ae17-d5f0-90f6-7834-4856ff7a1530
ms.date: 06/08/2017
---


# Selection.InsertColumnsRight Method (Word)

Inserts columns to the right of the current selection.


## Syntax

 _expression_ . **InsertColumnsRight**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Microsoft Word inserts as many columns as there are in the current selection.

To use this method, the current selection must be in a table.


## Example

This example selects the second column in the first table and inserts a new column to the right of it.


```vb
ActiveDocument.Tables(1).Columns(2).Select 
Selection.InsertColumnsRight
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

