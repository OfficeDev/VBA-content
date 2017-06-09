---
title: RotationEffect.By Property (PowerPoint)
keywords: vbapp10.chm661003
f1_keywords:
- vbapp10.chm661003
ms.prod: powerpoint
api_name:
- PowerPoint.RotationEffect.By
ms.assetid: 508d7a3e-ac92-af60-9f68-d394e78db363
ms.date: 06/08/2017
---


# RotationEffect.By Property (PowerPoint)

Sets or returns a  **Single** that represents the rotation of an object by the specified number of degrees; for example, a value of 180 means to rotate the object by 180 degrees. Read/write.


## Syntax

 _expression_. **By**

 _expression_ A variable that represents a **RotationEffect** object.


## Remarks

The specified object will be rotated with the center of the object remaining in the same position on the screen.

If both the  **By** and **[To](rotationeffect-to-property-powerpoint.md)** properties are set for a rotation effect, then the value of the **By** property is ignored.

Floating point numbers (for example, 55.5) are valid, but negative numbers are not.

Do not confuse this property with the  **ByX** or **ByY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

This example adds a rotation effect and changes its rotation.


```vb
Sub AddAndChangeRotationEffect()
    Dim effBlinds As Effect
    Dim tmlnShape As TimeLine
    Dim shpShape As Shape
    Dim animBehavior As AnimationBehavior
    Dim rtnEffect As RotationEffect

    'Sets shape, timing, and effect
    Set shpShape = ActivePresentation.Slides(1).Shapes(1)
    Set tmlnShape = ActivePresentation.Slides(1).TimeLine
    Set effBlinds = tmlnShape.MainSequence.AddEffect _
        (Shape:=shpShape, effectId:=msoAnimEffectBlinds)

    'Adds animation behavior and sets rotation effect
    Set animBehavior = tmlnShape.MainSequence(1).Behaviors _
        .Add(Type:=msoAnimTypeRotation)
    Set rtnEffect = animBehavior.RotationEffect

    rtnEffect.By = 270
End Sub
```


## See also


#### Concepts


[RotationEffect Object](rotationeffect-object-powerpoint.md)

