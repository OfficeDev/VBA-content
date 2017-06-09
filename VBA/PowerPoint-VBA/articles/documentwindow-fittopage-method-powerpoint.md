---
title: DocumentWindow.FitToPage Method (PowerPoint)
keywords: vbapp10.chm511015
f1_keywords:
- vbapp10.chm511015
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.FitToPage
ms.assetid: 91ea2102-df12-20fe-cd16-e664832f9eb5
ms.date: 06/08/2017
---


# DocumentWindow.FitToPage Method (PowerPoint)

Adjusts the size of the specified document window to accommodate the information that's currently displayed.


## Syntax

 _expression_. **FitToPage**

 _expression_ A variable that represents a **DocumentWindow** object.


## Example

This example exits the current slide show, sets the view in the active window to slide view, sets the zoom to 25 percent, and adjusts the size of the window to fit the slide displayed there.


```vb
Application.SlideShowWindows(1).View.Exit

With Application.ActiveWindow

    .ViewType = ppViewSlide

    .View.Zoom = 25

    .FitToPage

End With


```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

