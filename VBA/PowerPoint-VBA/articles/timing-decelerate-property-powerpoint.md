---
title: Timing.Decelerate Property (PowerPoint)
keywords: vbapp10.chm653011
f1_keywords:
- vbapp10.chm653011
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.Decelerate
ms.assetid: 3bf6fc1b-8f14-ef9a-cf70-69a93729f5bf
ms.date: 06/08/2017
---


# Timing.Decelerate Property (PowerPoint)

Sets or returns the percentageof the duration over which a timing deceleration should take place. Read/write.


## Syntax

 _expression_. **Decelerate**

 _expression_ A variable that represents a **Timing** object.


### Return Value

Single


## Remarks

For example, a value of 0.9 means that an deceleration should start at the default speed, and then start to slow down after the first ten percent of the animation. 


## Example

This example adds a shape and adds an animation that starts at the default speed and slows down after 70% of the animation has finished.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds rectangle and sets animation effect

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpRectangle, _
        effectId:=msoAnimEffectPathDiamond)

    'Slows the effect after seventy percent of the animation has finished

    With effDiamond.Timing
        .Decelerate = 0.3
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

