---
title: Selection.SelectColumn Method (Word)
keywords: vbawd10.chm158663172
f1_keywords:
- vbawd10.chm158663172
ms.prod: word
api_name:
- Word.Selection.SelectColumn
ms.assetid: a8e742df-0a8e-739d-e71a-da2536b6abec
ms.date: 06/08/2017
---


# Selection.SelectColumn Method (Word)

Selects the column that contains the insertion point, or selects all columns that contain the selection.


## Syntax

 _expression_ . **SelectColumn**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If the selection isn't in a table, an error occurs.


## Example

This example collapses the selection to the ending point and then selects the column that contains the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.Information(wdWithInTable) = True Then 
 Selection.SelectColumn 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

