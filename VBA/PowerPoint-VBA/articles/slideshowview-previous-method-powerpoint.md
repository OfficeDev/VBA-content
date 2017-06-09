---
title: SlideShowView.Previous Method (PowerPoint)
keywords: vbapp10.chm513020
f1_keywords:
- vbapp10.chm513020
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.Previous
ms.assetid: a53741b0-8325-696c-51e5-ffd3f9358ca8
ms.date: 06/08/2017
---


# SlideShowView.Previous Method (PowerPoint)

Shows the slide immediately preceding the slide that's currently displayed. 


## Syntax

 _expression_. **Previous**

 _expression_ A variable that represents a **SlideShowView** object.


## Remarks

If you are currently on the first slide in a kiosk slide show, the  **Previous** method takes you to the last slide in a slide show; otherwise, it has no effect if the first slide in the presentation is currently displayed. Use the **[View](slideshowwindow-view-property-powerpoint.md)** property of the **SlideShowWindow** object to return the **SlideShowView** object.


## Example

This example shows the slide immediately preceding the currently displayed slide on slide show window one.


```
SlideShowWindows(1).View.Previous
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

