---
title: Selection.Fields Property (Word)
keywords: vbawd10.chm158662720
f1_keywords:
- vbawd10.chm158662720
ms.prod: word
api_name:
- Word.Selection.Fields
ms.assetid: 15060502-c0cf-1c94-93ba-0db0bb045c66
ms.date: 06/08/2017
---


# Selection.Fields Property (Word)

Returns a read-only  **[Fields](fields-object-word.md)** collection that represents all the fields in the selection.


## Syntax

 _expression_ . **Fields**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example adds a DATE field at the insertion point.


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .Fields.Add Range:=Selection.Range, Type:=wdFieldDate 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

