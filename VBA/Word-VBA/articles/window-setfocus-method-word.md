---
title: Window.SetFocus Method (Word)
keywords: vbawd10.chm157417581
f1_keywords:
- vbawd10.chm157417581
ms.prod: word
api_name:
- Word.Window.SetFocus
ms.assetid: d6cf90ff-b62e-340d-140b-7d546d1f85a3
ms.date: 06/08/2017
---


# Window.SetFocus Method (Word)

Sets the focus of the specified document window to the body of an e-mail message.


## Syntax

 _expression_ . **SetFocus**

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


## Remarks

If the document isn't an e-mail message, this method has no effect.


## Example

This example makes the header of an e-mail message visible and sets the focus to the body of the message.


```vb
ActiveWindow.EnvelopeVisible = True 
ActiveWindow.SetFocus
```


## See also


#### Concepts


[Window Object](window-object-word.md)

