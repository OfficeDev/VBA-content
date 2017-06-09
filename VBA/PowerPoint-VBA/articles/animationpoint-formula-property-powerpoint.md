---
title: AnimationPoint.Formula Property (PowerPoint)
keywords: vbapp10.chm664006
f1_keywords:
- vbapp10.chm664006
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint.Formula
ms.assetid: 84ec9c9d-aa8b-faeb-8f51-a7fce91d709e
ms.date: 06/08/2017
---


# AnimationPoint.Formula Property (PowerPoint)

Returns or sets a  **String** that represents a formula to use for calculating an animation. Read/write.


## Syntax

 _expression_. **Formula**

 _expression_ A variable that represents a **AnimationPoint** object.


### Return Value

String


## Example

The following example adds a shape, and adds a three-second fill animation to that shape.


```vb
Sub AddShapeSetAnimFill()

    Dim effBlinds As Effect
    Dim shpRectangle As Shape
    Dim animBlinds As AnimationBehavior

    'Adds rectangle and sets animiation effect
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effBlinds = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectBlinds)

    'Sets the duration of the animation
    effBlinds.Timing.Duration = 3

    'Adds a behavior to the animation
    Set animBlinds = effBlinds.Behaviors.Add(msoAnimTypeProperty)

    'Sets the animation color effect and the formula to use
    With animBlinds.PropertyEffect
        .Property = msoAnimColor
        .Formula = RGB(Red:=255, Green:=255, Blue:=255)
    End With

End Sub
```


## See also


#### Concepts


[AnimationPoint Object](animationpoint-object-powerpoint.md)

