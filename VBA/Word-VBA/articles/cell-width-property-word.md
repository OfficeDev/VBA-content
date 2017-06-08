---
title: Cell.Width Property (Word)
keywords: vbawd10.chm156106758
f1_keywords:
- vbawd10.chm156106758
ms.prod: word
api_name:
- Word.Cell.Width
ms.assetid: 87c0422d-5f4f-44a3-902a-cb751b459ef9
ms.date: 06/08/2017
---


# Cell.Width Property (Word)

Returns or sets the width of a table cell, in points. Read/write  **Long** .


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example returns the width (in inches) of the cell that contains the insertion point.


```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox PointsToInches(Selection.Cells(1).Width) 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

