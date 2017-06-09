---
title: SlideShowView.GetClickIndex Method (PowerPoint)
keywords: vbapp10.chm513029
f1_keywords:
- vbapp10.chm513029
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.GetClickIndex
ms.assetid: 678feca3-79d4-e4e8-83aa-3484f5c099e9
ms.date: 06/08/2017
---


# SlideShowView.GetClickIndex Method (PowerPoint)

Returns the index number of the current mouse click for an animation that is actively playing on a slide or has just finished.


## Syntax

 _expression_. **GetClickIndex**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Long


## Remarks

Use the  **[GetClickCount](slideshowview-getclickcount-method-powerpoint.md)** method to return the number of mouse clicks that are defined for a slide.

If a slide has no animations or if a user has not advanced yet to an animation, the  **GetClickIndex** method returns 0. If a slide has an animation that runs automatically and the user moves to the previous page, the **GetClickIndex** method returns **msoClickStateBeforeAutomaticAnimations**.


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

