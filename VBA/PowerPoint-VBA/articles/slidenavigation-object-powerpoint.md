---
title: SlideNavigation Object (PowerPoint)
keywords: vbapp10.chm741000
f1_keywords:
- vbapp10.chm741000
ms.assetid: 3bb82afe-62a5-7e5a-597d-80f56f5cde4d
ms.date: 06/08/2017
ms.prod: powerpoint
---


# SlideNavigation Object (PowerPoint)

Represents the slide navigation screen in slide show view.


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


