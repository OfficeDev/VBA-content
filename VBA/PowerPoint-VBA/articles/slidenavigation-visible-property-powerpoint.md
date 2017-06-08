---
title: SlideNavigation.Visible Property (PowerPoint)
keywords: vbapp10.chm741002
f1_keywords:
- vbapp10.chm741002
ms.assetid: 76b526a1-2720-e6e6-9b94-07abed30e7ef
ms.date: 06/08/2017
ms.prod: powerpoint
---


# SlideNavigation.Visible Property (PowerPoint)

Specifies whether the slide navigation screen is displayed during a slide show. Read/write.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **SlideNavigation** object.


## Example

The following code sample starts a slide show from the active presentation and then makes the navigation screen visible.


```vb
Sub ShowSlideNavigation()

    ' Start the slide show.
    ActivePresentation.SlideShowSettings.Run
    
    ' Show the slide navigation screen.
    ActivePresentation.SlideShowWindow.SlideNavigation.Visible = True

End Sub
```


## Property value

 **BOOL**


