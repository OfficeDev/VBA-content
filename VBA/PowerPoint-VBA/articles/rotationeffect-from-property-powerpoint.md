---
title: RotationEffect.From Property (PowerPoint)
keywords: vbapp10.chm661004
f1_keywords:
- vbapp10.chm661004
ms.prod: powerpoint
api_name:
- PowerPoint.RotationEffect.From
ms.assetid: 9d5167f1-af74-3cfb-62b6-73afeadd10f8
ms.date: 06/08/2017
---


# RotationEffect.From Property (PowerPoint)

Sets or returns a  **Single** that represents the starting angle in degrees, specified relative to the screen (for example, 90 degrees is completely horizontal). Read/write.


## Syntax

 _expression_. **From**

 _expression_ A variable that represents a **RotationEffect** object.


## Remarks

Use this property in conjunction with the  **[To](rotationeffect-to-property-powerpoint.md)** property to transition from one rotation angle to another.

The default value is  **Empty** in which case the current position of the object is used.

Do not confuse this property with the  **FromX** or **FromY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

The following example adds a rotation effect and immediately changes its rotation angle.


```vb
Sub AddAndChangeRotationEffect()

    Dim effBlinds As Effect
    Dim tlnTiming As TimeLine
    Dim shpRectangle As Shape
    Dim animRotation As AnimationBehavior
    Dim rtnEffect As RotationEffect

    'Adds rectangle and sets effect and animation
    Set shpRectangle = ActivePresentation.Slides(1).Shapes_
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set tlnTiming = ActivePresentation.Slides(1).TimeLine

    Set effBlinds = tlnTiming.MainSequence.AddEffect(Shape:=shpRectangle, _
        effectId:=msoAnimEffectBlinds)

    Set animRotation = tlnTiming.MainSequence(1).Behaviors _
        .Add(Type:=msoAnimTypeRotation)

    Set rtnEffect = animRotation.RotationEffect

    'Sets the rotation effect starting and ending positions
    rtnEffect.From = 90
    rtnEffect.To = 270

End Sub
```


## See also


#### Concepts


[RotationEffect Object](rotationeffect-object-powerpoint.md)

