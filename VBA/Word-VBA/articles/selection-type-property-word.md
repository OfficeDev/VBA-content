---
title: Selection.Type Property (Word)
keywords: vbawd10.chm158662662
f1_keywords:
- vbawd10.chm158662662
ms.prod: word
api_name:
- Word.Selection.Type
ms.assetid: 75af6b1a-c9d3-e3ad-52a8-41d91c79b007
ms.date: 06/08/2017
---


# Selection.Type Property (Word)

Returns the selection type. Read-only  **[WdSelectionType](wdselectiontype-enumeration-word.md)** .


## Syntax

 _expression_ . **Type**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Example

This example formats the selection as engraved if the selection is not an insertion point.


```vb
If Selection.Type <> wdSelectionIP Then 
 Selection.Font.Engrave = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

