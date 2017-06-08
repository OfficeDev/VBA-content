---
title: TableOfFigures.RightAlignPageNumbers Property (Word)
keywords: vbawd10.chm153157635
f1_keywords:
- vbawd10.chm153157635
ms.prod: word
api_name:
- Word.TableOfFigures.RightAlignPageNumbers
ms.assetid: 0c9388b6-d6d7-9d41-547d-35d1345c1d38
ms.date: 06/08/2017
---


# TableOfFigures.RightAlignPageNumbers Property (Word)

 **True** if page numbers are aligned with the right margin in an table of figures. Read/write **Boolean** .


## Syntax

 _expression_ . **RightAlignPageNumbers**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Example

This example right-aligns page numbers for the first table of figures in the active document.


```vb
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 With ActiveDocument.TablesOfFigures(1) 
 .IncludePageNumbers = True 
 .RightAlignPageNumbers = True 
 End With 
End If
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

