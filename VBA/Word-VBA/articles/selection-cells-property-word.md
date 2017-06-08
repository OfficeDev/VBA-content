---
title: Selection.Cells Property (Word)
keywords: vbawd10.chm158662713
f1_keywords:
- vbawd10.chm158662713
ms.prod: word
api_name:
- Word.Selection.Cells
ms.assetid: 4b808b86-42ba-ccb4-b19a-87b134df3b79
ms.date: 06/08/2017
---


# Selection.Cells Property (Word)

Returns a  **[Cells](cells-object-word.md)** collection that represents the table cells in a selection. Read-only.


## Syntax

 _expression_ . **Cells**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the current cell's background color to red.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Shading.BackgroundPatternColorIndex = wdRed 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

