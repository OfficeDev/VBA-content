---
title: PageSetup.FirstSlideNumber Property (PowerPoint)
keywords: vbapp10.chm527003
f1_keywords:
- vbapp10.chm527003
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.FirstSlideNumber
ms.assetid: 277f613b-8c3a-d8bb-593c-a66ca41b4fa0
ms.date: 06/08/2017
---


# PageSetup.FirstSlideNumber Property (PowerPoint)

Returns or sets the slide number for the first slide in the presentation. Read/write.


## Syntax

 _expression_. **FirstSlideNumber**

 _expression_ A variable that represents a **PageSetup** object.


### Return Value

Long


## Remarks

The slide number is the actual number that will appear in the lower-right corner of the slide when you display slide numbers. This number is determined by the number (order) of the slide within the presentation (the  **SlideIndex** property value) and the starting slide number for the presentation (the **FirstSlideNumber** property value). The slide number will always be equal to the starting slide number + the slide index number - 1. The **SlideNumber** property returns the slide number.


## Example

This example shows how changing the first slide number in the active presentation affects the slide number of a specific slide.


```vb
With Application.ActivePresentation

    .PageSetup.FirstSlideNumber = 1  'starts slide numbering at 1

    MsgBox .Slides(2).SlideNumber    'returns 2

    .PageSetup.FirstSlideNumber = 10  'starts slide numbering at 10

    MsgBox .Slides(2).SlideNumber     'returns 11

End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

