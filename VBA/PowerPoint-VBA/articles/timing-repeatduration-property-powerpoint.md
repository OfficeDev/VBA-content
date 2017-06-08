---
title: Timing.RepeatDuration Property (PowerPoint)
keywords: vbapp10.chm653008
f1_keywords:
- vbapp10.chm653008
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.RepeatDuration
ms.assetid: 8c69f0a7-224a-db67-2a94-0237f55f184e
ms.date: 06/08/2017
---


# Timing.RepeatDuration Property (PowerPoint)

Sets or returns how long repeated animations should last, in seconds. Read/write.


## Syntax

 _expression_. **RepeatDuration**

 _expression_ A variable that represents a **Timing** object.


### Return Value

Single


## Remarks

An animation will stop at the end of its time sequence or the value of the  **RepeatDuration** property, whichever is shorter.


## Example

This examples adds a shape and an animation to it, then repeats the animation ten times. However, after five seconds, the animation will be cut off, even though the animation is dimensioned for a 20-second timeline (if the  **Duration** property is not specified, an animation defaults to two seconds).


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds new shape and sets animation effect
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    'Sets repeat duration and number of times to repeat animation
    With effDiamond.Timing
        .RepeatDuration = 5
        .RepeatCount = 10
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

