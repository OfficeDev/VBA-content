---
title: AnimationPoint.Time Property (PowerPoint)
keywords: vbapp10.chm664004
f1_keywords:
- vbapp10.chm664004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint.Time
ms.assetid: 19df62b1-b898-fdba-d5e4-86ac5a68cecf
ms.date: 06/08/2017
---


# AnimationPoint.Time Property (PowerPoint)

Sets or returns the time at a given animation point. Read/write.


## Syntax

 _expression_. **Time**

 _expression_ A variable that represents a **SlideShowTransition** object.


### Return Value

Single


## Remarks

The value of the  **Time** property can be any floating-point value between 0 and 1, representing a percentage of the entire timeline from 0% to 100%. For example, a value of 0.2 would correspond to a point in time at 20% of the entire timeline duration from left to right.


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

