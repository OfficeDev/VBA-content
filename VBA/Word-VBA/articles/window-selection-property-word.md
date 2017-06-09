---
title: Window.Selection Property (Word)
keywords: vbawd10.chm157417476
f1_keywords:
- vbawd10.chm157417476
ms.prod: word
api_name:
- Word.Window.Selection
ms.assetid: 0e6812cd-8b8a-edaf-cf72-cf899c50f92a
ms.date: 06/08/2017
---


# Window.Selection Property (Word)

Returns the  **Selection** object that represents a selected range or the insertion point. Read-only.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example copies the selection from window one to the next window.


```vb
If Windows.Count >= 2 Then 
 Windows(1).Selection.Copy 
 Windows(1).Next.Activate 
 Selection.Paste 
End If
```


## See also


#### Concepts


[Window Object](window-object-word.md)

