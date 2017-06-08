---
title: AnimationPoint.Value Property (PowerPoint)
keywords: vbapp10.chm664005
f1_keywords:
- vbapp10.chm664005
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint.Value
ms.assetid: f16879c0-25cc-46fa-cfd3-7a6a770be371
ms.date: 06/08/2017
---


# AnimationPoint.Value Property (PowerPoint)

Sets or returns the value of a property for an animation point. Read/write.


## Syntax

 _expression_. **Value**

 _expression_ A variable that represents an **AnimationPoint** object.


### Return Value

Variant


## Example

This example inserts three fill color animation points in the main sequence animation timeline on the first slide.


```vb
Sub BuildTimeLine()

    Dim shpFirst As Shape
    Dim effMain As Effect
    Dim tmlMain As TimeLine
    Dim aniBhvr As AnimationBehavior
    Dim aniPoint As AnimationPoint

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)
    Set tmlMain = ActivePresentation.Slides(1).TimeLine
    Set effMain = tmlMain.MainSequence.AddEffect(Shape:=shpFirst, _
        EffectId:=msoAnimEffectBlinds)

    Set aniBhvr = tmlMain.MainSequence(1).Behaviors.Add _
        (Type:=msoAnimTypeProperty)

    With aniBhvr.PropertyEffect
        .Property = msoAnimShapeFillColor
        Set aniPoint = .Points.Add
        aniPoint.Time = 0.2
        aniPoint.Value = RGB(0, 0, 0)
        Set aniPoint = .Points.Add
        aniPoint.Time = 0.5
        aniPoint.Value = RGB(0, 255, 0)
        Set aniPoint = .Points.Add
        aniPoint.Time = 1
        aniPoint.Value = RGB(0, 255, 255)
    End With

End Sub
```


## See also


#### Concepts


[AnimationPoint Object](animationpoint-object-powerpoint.md)

