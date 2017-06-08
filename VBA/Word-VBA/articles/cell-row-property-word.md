---
title: Cell.Row Property (Word)
keywords: vbawd10.chm156106854
f1_keywords:
- vbawd10.chm156106854
ms.prod: word
api_name:
- Word.Cell.Row
ms.assetid: b395a2f8-2eb4-1443-1298-56e3d3ad068b
ms.date: 06/08/2017
---


# Cell.Row Property (Word)

Returns a  **[Row](row-object-word.md)** object that represents the row containing the specified cell.


## Syntax

 _expression_ . **Row**

 _expression_ An expression that returns a **[Cell](cell-object-word.md)** object.


## Example

This example applies shading to the table row that contains the insertion point.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Row.Shading.Texture = wdTexture10Percent 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

