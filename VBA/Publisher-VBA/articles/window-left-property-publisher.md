---
title: Window.Left Property (Publisher)
keywords: vbapb10.chm262149
f1_keywords:
- vbapb10.chm262149
ms.prod: publisher
api_name:
- Publisher.Window.Left
ms.assetid: 8d61331a-a70f-4a8a-8dc7-12d93ec51bfc
ms.date: 06/08/2017
---


# Window.Left Property (Publisher)

Returns or sets a  **Long** indicating the position (in points) of the left edge of the application window relative to the left edge of the screen. Read/write.


## Syntax

 _expression_. **Left**

 _expression_A variable that represents a  **Window** object.


## Example

This example sets the horizontal position of the active window to 100 points.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = pbWindowStateNormal 
 .Left = 100 
 .Top = 0 
End With
```


