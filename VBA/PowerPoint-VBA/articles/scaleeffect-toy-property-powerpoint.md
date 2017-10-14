---
title: ScaleEffect.ToY Property (PowerPoint)
keywords: vbapp10.chm660008
f1_keywords:
- vbapp10.chm660008
ms.prod: powerpoint
api_name:
- PowerPoint.ScaleEffect.ToY
ms.assetid: 987b5456-ea44-0fa8-3f2b-7bc0c9dccbe0
ms.date: 06/08/2017
---


# ScaleEffect.ToY Property (PowerPoint)

Returns or sets a  **Single** that represents the ending height of a **[ScaleEffect](scaleeffect-object-powerpoint.md)** object, specified as a percentage of the screen width. Read/write.


## Syntax

 _expression_. **ToY**

 _expression_ A variable that represents a **ScaleEffect** object.


### Return Value

Single


## Remarks

The default value of this property is  **Empty**, in which case the current position of the object is used.

Use this property in conjunction with the  **FromY** property to resize or jump from one position to another.

Do not confuse this property with the  **To** property of the **[ColorEffect](coloreffect-object-powerpoint.md)**, **[RotationEffect](rotationeffect-object-powerpoint.md)**, or **[PropertyEffect](propertyeffect-object-powerpoint.md)** objects, which is used to set or change colors, rotations, or other properties of an animation behavior, respectively.


## Example

The following example adds an animation path and sets the starting and ending horizontal and vertical positions.


```vb
Sub AddMotionPath()

    Dim effCustom As Effect
    Dim animMotion As AnimationBehavior
    Dim shpRectangle As Shape

    'Adds shape and sets effect and animation properties
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effCustom = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectCustom)
 
    Set animMotion = effCustom.Behaviors.Add(msoAnimTypeMotion)

    'Sets starting and ending horizontal and vertical positions
    With animMotion.MotionEffect
        .FromX = 0
        .FromY = 0
        .ToX = 50
        .ToY = 50
    End With

End Sub
```


## See also


#### Concepts


[ScaleEffect Object](scaleeffect-object-powerpoint.md)

