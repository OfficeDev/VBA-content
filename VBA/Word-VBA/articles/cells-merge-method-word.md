---
title: Cells.Merge Method (Word)
keywords: vbawd10.chm155844812
f1_keywords:
- vbawd10.chm155844812
ms.prod: word
api_name:
- Word.Cells.Merge
ms.assetid: 064d405e-00a1-205a-184c-4f46ab463a63
ms.date: 06/08/2017
---


# Cells.Merge Method (Word)

Merges the specified table cells with one another. The result is a single table cell.


## Syntax

 _expression_ . **Merge**

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


## Example

This example merges the cells in row one of the selection into a single cell and then applies shading to the row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Set myrow = Selection.Rows(1) 
 myrow.Cells.Merge 
 myrow.Shading.Texture = wdTexture10Percent 
End If
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

