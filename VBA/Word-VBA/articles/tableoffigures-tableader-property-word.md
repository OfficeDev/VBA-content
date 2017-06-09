---
title: TableOfFigures.TabLeader Property (Word)
keywords: vbawd10.chm153157644
f1_keywords:
- vbawd10.chm153157644
ms.prod: word
api_name:
- Word.TableOfFigures.TabLeader
ms.assetid: c806034e-f226-0be8-aa29-25f9b85b2a39
ms.date: 06/08/2017
---


# TableOfFigures.TabLeader Property (Word)

Returns or sets the character between entries and their page numbers in an table of figures. Read/write  **[WdTabLeader](wdtableader-enumeration-word.md)** .


## Syntax

 _expression_ . **TabLeader**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Example

This example formats the tables of firgures in Sales.doc to use a dotted tab leader.


```vb
For Each aTOF In Documents("Sales.doc").TablesOfFigures 
 aTOF.TabLeader = wdTabLeaderDots 
Next aTOF
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

