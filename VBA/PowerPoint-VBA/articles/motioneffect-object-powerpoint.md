---
title: MotionEffect Object (PowerPoint)
keywords: vbapp10.chm658000
f1_keywords:
- vbapp10.chm658000
ms.prod: powerpoint
api_name:
- PowerPoint.MotionEffect
ms.assetid: 77a34f68-8806-22b8-149f-c28e0457e7e9
ms.date: 06/08/2017
---


# MotionEffect Object (PowerPoint)

Represents a motion effect for an  **AnimationBehavior** object.


## Example

Use the [MotionEffect](animationbehavior-motioneffect-property-powerpoint.md)propety of the  **AnimationBehavior** object to return a **MotionEffect** object. The following example refers to the motion effect for a given animation behavior.


```vb
ActivePresentation.Slides(1).TimeLine.MainSequence.Item.Behaviors(1).MotionEffect
```

Use the [ByX](motioneffect-byx-property-powerpoint.md), [ByY](motioneffect-byy-property-powerpoint.md), [FromX](motioneffect-fromx-property-powerpoint.md), [FromY](motioneffect-fromy-property-powerpoint.md), [ToX](motioneffect-tox-property-powerpoint.md), and [ToY](motioneffect-toy-property-powerpoint.md)properties of the  **MotionEffect** object to construct a motion path. The **ToY** and **ToX** properties are in percentage, where **ToX** = 1.0 means 100% of slide width and **ToY** = 1.0 means 100% of slide height. The following example adds a shape to the first slide and creates a motion path.




```vb
Sub AddMotionPath()

    Dim shpNew As Shape
    Dim effNew As Effect
    Dim aniMotion As AnimationBehavior

    Set shpNew = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShape5pointStar, Left:=0, _
        Top:=0, Width:=100, Height:=100)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpNew, effectId:=msoAnimEffectCustom, _
        Trigger:=msoAnimTriggerWithPrevious)

    Set aniMotion = effNew.Behaviors.Add(msoAnimTypeMotion)

    With aniMotion.MotionEffect
        .FromX = 0
        .FromY = 0
        .ToX = .5
        .ToY = .5
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

