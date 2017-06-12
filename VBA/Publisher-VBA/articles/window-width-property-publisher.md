---
title: Window.Width Property (Publisher)
keywords: vbapb10.chm262150
f1_keywords:
- vbapb10.chm262150
ms.prod: publisher
api_name:
- Publisher.Window.Width
ms.assetid: 762df30a-7fdd-8f95-f64b-eae57e7c02fe
ms.date: 06/08/2017
---


# Window.Width Property (Publisher)

Returns or sets a  **Long** that represents the width (in points) of the window. Read/write.


## Syntax

 _expression_. **Width**

 _expression_A variable that represents a  **Window** object.


## Example

This example sets the height and width of the active window if the window is neither maximized nor minimized.


```vb
Sub SetWindowHeight() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Height = InchesToPoints(5) 
 .Width = InchesToPoints(5) 
 End If 
 End With 
End Sub
```


