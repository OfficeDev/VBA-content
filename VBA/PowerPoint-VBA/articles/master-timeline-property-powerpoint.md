---
title: Master.TimeLine Property (PowerPoint)
keywords: vbapp10.chm533015
f1_keywords:
- vbapp10.chm533015
ms.prod: powerpoint
api_name:
- PowerPoint.Master.TimeLine
ms.assetid: f57756b5-9b13-336b-0d5c-00161590ba03
ms.date: 06/08/2017
---


# Master.TimeLine Property (PowerPoint)

Returns a  **[TimeLine](timeline-object-powerpoint.md)** object that represents the animation timeline for the slide. Read-only.


## Syntax

 _expression_. **TimeLine**

 _expression_ A variable that represents a **Master** object.


### Return Value

TimeLine


## Example

The following example adds a bouncing animation to the first shape on the first slide.


```vb
Sub NewTimeLineEffect()

    Dim sldFirst As Slide
    Dim shpFirst As Shape

    Set sldFirst = ActivePresentation.Slides(1)
    Set shpFirst = sldFirst.Shapes(1)

    sldFirst.TimeLine.MainSequence.AddEffect _
        Shape:=shpFirst, EffectId:=msoAnimEffectBounce

End Sub
```


## See also


#### Concepts


[Master Object](master-object-powerpoint.md)

