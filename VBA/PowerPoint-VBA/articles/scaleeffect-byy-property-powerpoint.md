---
title: ScaleEffect.ByY Property (PowerPoint)
keywords: vbapp10.chm660004
f1_keywords:
- vbapp10.chm660004
ms.prod: powerpoint
api_name:
- PowerPoint.ScaleEffect.ByY
ms.assetid: c77a59cb-dc68-120b-8750-3088ccb12d73
ms.date: 06/08/2017
---


# ScaleEffect.ByY Property (PowerPoint)

Sets or returns a  **Single** that represents scaling or moving an object vertically by a specified percentage of the screen width, depending on whether it is used in conjunction with a **[ScaleEffect](scaleeffect-object-powerpoint.md)** or **[MotionEffect](motioneffect-object-powerpoint.md)** object, respectively. Read/write.


## Syntax

 _expression_. **ByY**

 _expression_ A variable that represents a **ScaleEffect** object.


### Return Value

Single


## Remarks

Negative numbers move the object horizontally to the left. Floating point numbers (for example, 55.5) are allowed.

To scale or move an object horizontally, use the  **ByX** property.

If both the  **ByX** and **ByY** properties are set, then the object is scaled or moves both horizontally and vertically.

Do not confuse this property with the  **By** property of the **[ColorEffect](coloreffect-object-powerpoint.md)**, **[RotationEffect](rotationeffect-object-powerpoint.md)**, or **[PropertyEffect](propertyeffect-object-powerpoint.md)** objects, which is used to set colors, rotations, or other properties of an animation behavior, respectively.


## Example

The following example adds an animation path; then sets the horizontal and vertical movement of the shape.


```vb
Sub AddMotionPath()

    Dim effCustom As Effect
    Dim animBehavior As AnimationBehavior
    Dim shpRectangle As Shape

    'Adds rectangle and sets effect and animation
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=300, _
        Top:=300, Width:=300, Height:=150)

    Set effCustom = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpRectangle, _
         effectId:=msoAnimEffectCustom)

    Set animBehavior = effCustom.Behaviors.Add(msoAnimTypeMotion)

    'Specifies animation motion
    With animBehavior.MotionEffect
        .ByX = 50
        .ByY = 50
    End With

End Sub
```


## See also


#### Concepts


[ScaleEffect Object](scaleeffect-object-powerpoint.md)

