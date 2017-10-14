---
title: ScaleEffect Object (PowerPoint)
keywords: vbapp10.chm660000
f1_keywords:
- vbapp10.chm660000
ms.prod: powerpoint
api_name:
- PowerPoint.ScaleEffect
ms.assetid: cb7c296e-a9ea-4ed6-87e0-a5d603da4f9f
ms.date: 06/08/2017
---


# ScaleEffect Object (PowerPoint)

Represents a scaling effect for an  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object.


## Example

Use the [ScaleEffect](animationbehavior-scaleeffect-property-powerpoint.md)property of the  **AnimationBehavior** object to return a **ScaleEffect** object. The following example refers to the scale effect for a given animation behavior.


```vb
ActivePresentation.Slides(1).TimeLine.MainSequence.Item.Behaviors(1).ScaleEffect
```

Use the [ByX](scaleeffect-byx-property-powerpoint.md), [ByY](scaleeffect-byy-property-powerpoint.md), [FromX](scaleeffect-fromx-property-powerpoint.md), [FromY](scaleeffect-fromy-property-powerpoint.md), [ToX](scaleeffect-tox-property-powerpoint.md), and [ToY](scaleeffect-toy-property-powerpoint.md)properties of the  **ScaleEffect** object to manipulate an object's scale. This example scales the first shape on the first slide starting at zero increasing in size until it reaches 100 percent of its original size. This example assumes that there is a shape on the first slide.




```vb
Sub ChangeScale()

    Dim shpFirst As Shape
    Dim effNew As Effect
    Dim aniScale As AnimationBehavior

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)
    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpFirst, effectId:=msoAnimEffectCustom)

    Set aniScale = effNew.Behaviors.Add(msoAnimTypeScale)
    With aniScale.ScaleEffect
        'Starting size
        .FromX = 0
        .FromY = 0

        'Size after scale effect
        .ToX = 100
        .ToY = 100
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

