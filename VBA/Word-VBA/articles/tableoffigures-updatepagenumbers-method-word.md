---
title: TableOfFigures.UpdatePageNumbers Method (Word)
keywords: vbawd10.chm153157733
f1_keywords:
- vbawd10.chm153157733
ms.prod: word
api_name:
- Word.TableOfFigures.UpdatePageNumbers
ms.assetid: d6817167-916d-81f0-c507-16492819b0f3
ms.date: 06/08/2017
---


# TableOfFigures.UpdatePageNumbers Method (Word)

Updates the page numbers for items in a table of figures.


## Syntax

 _expression_ . **UpdatePageNumbers**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Example

This example updates all tables of figures in Sales.doc.


```vb
Dim tofLoop As TableOfFigures 
 
For Each tofLoop In Documents("Sales.doc").TablesOfFigures 
 tofLoop.UpdatePageNumbers 
Next tofLoop
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

