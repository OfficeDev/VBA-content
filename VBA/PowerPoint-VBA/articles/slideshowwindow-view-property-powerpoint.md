---
title: SlideShowWindow.View Property (PowerPoint)
keywords: vbapp10.chm507003
f1_keywords:
- vbapp10.chm507003
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow.View
ms.assetid: ebf565af-fc90-ab1b-0e05-6dcb90a7c2d2
ms.date: 06/08/2017
---


# SlideShowWindow.View Property (PowerPoint)

Returns a  **[SlideShowView](slideshowview-object-powerpoint.md)** object. Read-only.


## Syntax

 _expression_. **View**

 _expression_ A variable that represents a **SlideShowWindow** object.


### Return Value

SlideShowView


## Example

This example uses the  **View** property to exit the current slide show, sets the view in the active window to slide view, and then displays slide three.


```vb
Application.SlideShowWindows(1).View.Exit

With Application.ActiveWindow

    .ViewType = ppViewSlide

    .View.GotoSlide 3

End With
```


## See also


#### Concepts


[SlideShowWindow Object](slideshowwindow-object-powerpoint.md)

