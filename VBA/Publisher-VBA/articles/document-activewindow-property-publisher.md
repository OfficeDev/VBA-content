---
title: Document.ActiveWindow Property (Publisher)
keywords: vbapb10.chm196611
f1_keywords:
- vbapb10.chm196611
ms.prod: publisher
api_name:
- Publisher.Document.ActiveWindow
ms.assetid: 0d00a8fa-aef2-43df-3c54-0cca804b7eee
ms.date: 06/08/2017
---


# Document.ActiveWindow Property (Publisher)

Returns a  **[Window](window-object-publisher.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.


## Syntax

 _expression_. **ActiveWindow**

 _expression_A variable that represents a  **Document** object.


## Example

This example displays the active window's caption.


```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```


