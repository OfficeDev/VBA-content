---
title: AnimationBehavior.RotationEffect Property (PowerPoint)
keywords: vbapp10.chm657009
f1_keywords:
- vbapp10.chm657009
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.RotationEffect
ms.assetid: 46983cf0-0977-41ec-6264-958216ee44dc
ms.date: 06/08/2017
---


# AnimationBehavior.RotationEffect Property (PowerPoint)

Returns a  **[RotationEffect](rotationeffect-object-powerpoint.md)** object for an animation behavior. Read-only.


## Syntax

 _expression_. **RotationEffect**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

RotationEffect


## Example

The following example adds a new shape to the first slide and sets the rotation animation behavior.


```vb
Sub AddRotation()

    Dim shpNew As Shape
    Dim effNew As Effect
    Dim aniNew As AnimationBehavior

    Set shpNew = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShape5pointStar, Left:=0, _
        Top:=0, Width:=100, Height:=100)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpNew, effectId:=msoAnimEffectCustom)

    Set aniNew = effNew.Behaviors.Add(msoAnimTypeRotation)

    With aniNew.RotationEffect
        'Rotate 270 degrees from current position
        .By = 270
    End With

End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

