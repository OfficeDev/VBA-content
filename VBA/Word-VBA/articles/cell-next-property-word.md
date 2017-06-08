---
title: Cell.Next Property (Word)
keywords: vbawd10.chm156106855
f1_keywords:
- vbawd10.chm156106855
ms.prod: word
api_name:
- Word.Cell.Next
ms.assetid: b4171c7c-6703-9cdf-a964-09e32874fbb6
ms.date: 06/08/2017
---


# Cell.Next Property (Word)

Returns a  **Cell** object that represents the next table cell in the **Cells** collection. Read-only.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

If the selection is in a table, this example selects the contents of the next table cell.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Next.Select 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

