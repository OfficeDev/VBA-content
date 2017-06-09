---
title: Window.Activate Method (Word)
keywords: vbawd10.chm157417572
f1_keywords:
- vbawd10.chm157417572
ms.prod: word
api_name:
- Word.Window.Activate
ms.assetid: d068e7a1-edb8-b244-a315-be1f92471f4c
ms.date: 06/08/2017
---


# Window.Activate Method (Word)

Activates the specified window.


## Syntax

 _expression_ . **Activate**

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example activates the next window in the Windows collection.


```vb
Sub NextWindow() 
 'Two or more documents must be open for this statement to execute. 
 ActiveDocument.ActiveWindow.Next.Activate 
End Sub
```


## See also


#### Concepts


[Window Object](window-object-word.md)

