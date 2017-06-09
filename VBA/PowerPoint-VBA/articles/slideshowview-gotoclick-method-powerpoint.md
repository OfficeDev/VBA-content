---
title: SlideShowView.GotoClick Method (PowerPoint)
keywords: vbapp10.chm513028
f1_keywords:
- vbapp10.chm513028
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.GotoClick
ms.assetid: b41dec86-96a9-447a-5895-0b28fc4bd6b2
ms.date: 06/08/2017
---


# SlideShowView.GotoClick Method (PowerPoint)

Plays an animation associated with a specified mouse click and any animations that follow on the slide.


## Syntax

 _expression_. **GotoClick**( **_Index_** )

 _expression_ A variable that represents a **SlideShowView** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number of the mouse click that initiates an animation. |

## Remarks

Use the  **[GetClickIndex](slideshowview-getclickindex-method-powerpoint.md)** method to return the index number of the current mouse click for an animation that is actively playing on a slide or has just finished.

Specifying a value of 0 for Index plays animations beginning at the point just before any animations that run automatically. Specifying a value of  **msoClickStateBeforeAutomaticAnimations** for Index moves to the point just before any animations that run automatically, and then pauses. Specifying an value of **msoClickStateAfterAllAnimations** for Index moves to the point after all animations.


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

