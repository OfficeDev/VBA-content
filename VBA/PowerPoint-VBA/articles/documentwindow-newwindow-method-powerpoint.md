---
title: DocumentWindow.NewWindow Method (PowerPoint)
keywords: vbapp10.chm511019
f1_keywords:
- vbapp10.chm511019
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.NewWindow
ms.assetid: 1c9f4e37-4e40-8d0b-246b-f9897ad9a56a
ms.date: 06/08/2017
---


# DocumentWindow.NewWindow Method (PowerPoint)

Opens a new window that contains the same document that is displayed in the specified window. Returns a  **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the new window.


## Syntax

 _expression_. **NewWindow**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

DocumentWindow


## Example

This example creates a new window that contains the contents of the active window (thereby activating the new window) and then switches back to the first window.


```vb
Set oldW = Application.ActiveWindow

Set newW = oldW.NewWindow

oldW.Activate
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


