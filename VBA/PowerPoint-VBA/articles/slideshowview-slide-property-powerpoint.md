---
title: SlideShowView.Slide Property (PowerPoint)
keywords: vbapp10.chm513004
f1_keywords:
- vbapp10.chm513004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.Slide
ms.assetid: 4fdee96b-9b0d-64ba-19de-b810bf07987b
ms.date: 06/08/2017
---


# SlideShowView.Slide Property (PowerPoint)

Returns a  **[Slide](slide-object-powerpoint.md)** object that represents the slide that's currently displayed in the specified slide show window view. Read-only.


## Syntax

 _expression_. **Slide**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Slide


## Remarks

If the currently displayed slide is from an embedded presentation, you can use the  **[Parent](slide-parent-property-powerpoint.md)** property of the **Slide** object returned by the **Slide** property to return the embedded presentation that contains the slide. (The **[Presentation](slideshowwindow-presentation-property-powerpoint.md)** property of the **SlideShowWindow** object or **DocumentWindow** object returns the presentation from which the window was created, not the embedded presentation.)


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

