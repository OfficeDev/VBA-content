---
title: Selection.IsEndOfRowMark Property (Word)
keywords: vbawd10.chm158662963
f1_keywords:
- vbawd10.chm158662963
ms.prod: word
api_name:
- Word.Selection.IsEndOfRowMark
ms.assetid: 0729a8f2-628c-902b-fca1-488742234873
ms.date: 06/08/2017
---


# Selection.IsEndOfRowMark Property (Word)

 **True** if the specified selection or range is collapsed and is located at the end-of-row mark in a table. Read-only **Boolean** .


## Syntax

 _expression_ . **IsEndOfRowMark**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This property is the equivalent of the following expression:


```
Selection.Information(wdAtEndOfRowMarker)
```


## Example

This example collapses the selection and selects the current row if the insertion point is at the end of the row (just before the end-of-row mark).


```vb
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.IsEndOfRowMark = True Then 
 Selection.Rows(1).Select 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

