---
title: DocumentWindow.Presentation Property (PowerPoint)
keywords: vbapp10.chm511005
f1_keywords:
- vbapp10.chm511005
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.Presentation
ms.assetid: f009e2c3-aa08-09f0-c879-a25b8d1e0405
ms.date: 06/08/2017
---


# DocumentWindow.Presentation Property (PowerPoint)

Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the presentation in which the specified document window or slide show window was created. Read-only.


## Syntax

 _expression_. **Presentation**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

Presentation


## Remarks

If the slide that's currently displayed in document window one is from an embedded presentation,  `Windows(1).View.Slide.Parent` returns the embedded presentation, and `Windows(1).Presentation` returns the presentation in which document window one was created.

If the slide that's currently displayed in slide show window one is from an embedded presentation,  `SlideShowWindows(1).View.Slide.Parent` returns the embedded presentation, and `SlideShowWindows(1).Presentation` returns the presentation in which the slide show was started.


## Example

This example continues the slide numbering for the presentation in window one into the slide numbering for the presentation in window two.


```
firstPresSlides = Windows(1).Presentation.Slides.Count

Windows(2).Presentation.PageSetup _
    .FirstSlideNumber = firstPresSlides + 1
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


