---
title: Selection.SelectRow Method (Word)
keywords: vbawd10.chm158663181
f1_keywords:
- vbawd10.chm158663181
ms.prod: word
api_name:
- Word.Selection.SelectRow
ms.assetid: 0d821d49-2829-2469-4742-0355440e4775
ms.date: 06/08/2017
---


# Selection.SelectRow Method (Word)

Selects the row that contains the insertion point, or selects all rows that contain the selection.


## Syntax

 _expression_ . **SelectRow**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If the selection isn't in a table, an error occurs.


## Example

This example collapses the selection to the starting point and then selects the column that contains the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.SelectRow 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

