---
title: SlideShowView.LastSlideViewed Property (PowerPoint)
keywords: vbapp10.chm513010
f1_keywords:
- vbapp10.chm513010
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.LastSlideViewed
ms.assetid: 47647e03-d898-47b5-cb50-79f3e368b56f
ms.date: 06/08/2017
---


# SlideShowView.LastSlideViewed Property (PowerPoint)

Returns a  **[Slide](slide-object-powerpoint.md)** object that represents the slide viewed immediately before the current slide in the specified slide show view.


## Syntax

 _expression_. **LastSlideViewed**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Slide


## Example

This example takes you to the slide viewed immediately before the current slide in slide show window one.


```vb
With SlideShowWindows(1).View

    .GotoSlide (.LastSlideViewed.SlideIndex)

End With
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

