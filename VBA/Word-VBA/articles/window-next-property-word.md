---
title: Window.Next Property (Word)
keywords: vbawd10.chm157417488
f1_keywords:
- vbawd10.chm157417488
ms.prod: word
api_name:
- Word.Window.Next
ms.assetid: 28587dfe-dd49-88b7-0261-b4e42a12eeac
ms.date: 06/08/2017
---


# Window.Next Property (Word)

Returns the next document window in the collection of open document windows. Read-only.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example activates the next window.


```vb
If Windows.Count > 1 Then ActiveDocument.ActiveWindow.Next.Activate
```


## See also


#### Concepts


[Window Object](window-object-word.md)

