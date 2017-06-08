---
title: RotationEffect.To Property (PowerPoint)
keywords: vbapp10.chm661005
f1_keywords:
- vbapp10.chm661005
ms.prod: powerpoint
api_name:
- PowerPoint.RotationEffect.To
ms.assetid: 9630d2d6-818c-d86b-dbd7-54b3b2b13ad2
ms.date: 06/08/2017
---


# RotationEffect.To Property (PowerPoint)

Sets or returns a  **Single** that represents the ending rotation of an object in degrees, specified relative to the screen (for example, 90 degrees is completely horizontal). Read/write.


## Syntax

 _expression_. **To**

 _expression_ A variable that represents a **RotationEffect** object.


### Return Value

Single


## Remarks

Use this property in conjunction with the  **From** property to transition from one rotation angle to another.

The default value is  **Empty** in which case the current position of the object is used.

Do not confuse this property with the  **ToX** or **ToY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

The following example adds a rotation effect and immediately changes its rotation angle from 90 degrees to 270 degrees.


```vb
Sub AddAndChangeRotationEffect()

    Dim effBlinds As Effect
    Dim tmlTiming As TimeLine
    Dim shpRectangle As Shape
    Dim animColor As AnimationBehavior
    Dim rtnEffect As RotationEffect

    Set shpRectangle = ActivePresentation.Slides(1).Shapes(1)
    Set tmlTiming = ActivePresentation.Slides(1).TimeLine
    Set effBlinds = tmlTiming.MainSequence.AddEffect(Shape:=shpRectangle, _
        effectId:=msoAnimEffectBlinds)
    Set animColor = tmlTiming.MainSequence(1).Behaviors.Add(Type:=msoAnimTypeRotation)
    Set rtnEffect = animColor.RotationEffect
    rtnEffect.From = 90
    rtnEffect.To = 270

End Sub
```


## See also


#### Concepts


[RotationEffect Object](rotationeffect-object-powerpoint.md)

