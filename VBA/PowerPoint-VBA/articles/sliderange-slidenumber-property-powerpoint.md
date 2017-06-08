---
title: SlideRange.SlideNumber Property (PowerPoint)
keywords: vbapp10.chm532019
f1_keywords:
- vbapp10.chm532019
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.SlideNumber
ms.assetid: 2fee7966-26b8-4ed1-adb4-cff9ce6ba37f
ms.date: 06/08/2017
---


# SlideRange.SlideNumber Property (PowerPoint)

Returns the slide number. Read-only.


## Syntax

 _expression_. **SlideNumber**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

Integer


## Remarks

The  **SlideNumber** property of a **Slide** object is the actual number that appears in the lower-right corner of the slide when you display slide numbers. This number is determined by the number of the slide within the presentation (the **[SlideIndex](sliderange-slideindex-property-powerpoint.md)** property value) and the starting slide number for the presentation (the **[FirstSlideNumber](pagesetup-firstslidenumber-property-powerpoint.md)** property value). The slide number is always equal to the starting slide number + the slide index number - 1.


## Example

This example shows how changing the first slide number affects the slide number of a specific slide.


```vb
With Application.ActivePresentation

    .PageSetup.FirstSlideNumber = 1   'starts slide numbering at 1

    MsgBox .Slides(2).SlideNumber     'returns 2



    .PageSetup.FirstSlideNumber = 10 'starts slide numbering at 10

    MsgBox .Slides(2).SlideNumber     'returns 11

End With
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

