---
title: Window.Height Property (Publisher)
keywords: vbapb10.chm262151
f1_keywords:
- vbapb10.chm262151
ms.prod: publisher
api_name:
- Publisher.Window.Height
ms.assetid: 3d47bb99-bab7-b5aa-c834-04bcd6e8b151
ms.date: 06/08/2017
---


# Window.Height Property (Publisher)

Returns or sets a  **Long** that represents the height (in points) of the window. Read/write.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **Window** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


## Example

This example sets the height and width of the active window if the window is neither maximized nor minimized.


```vb
Sub SetWindowHeight() 
 With ActiveWindow 
 If .WindowState <> pbWindowStateNormal Then 
 .WindowState = pbWindowStateNormal 
 .Height = InchesToPoints(5) 
 .Width = InchesToPoints(5) 
 End If 
 End With 
End Sub
```


