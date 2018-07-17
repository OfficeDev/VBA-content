---
title: FilterEffect Object (PowerPoint)
keywords: vbapp10.chm669000
f1_keywords:
- vbapp10.chm669000
ms.prod: powerpoint
api_name:
- PowerPoint.FilterEffect
ms.assetid: f61235e0-5ddc-536e-1ac1-92b8b519f130
ms.date: 06/08/2017
---


# FilterEffect Object (PowerPoint)

Represents a filter effect for an animation behavior.


## Remarks

Use the  **FilterEffect** property of the **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object to return a **FilterEffect** object. Filter effects can be changed using the **FilterEffect** object's **Reveal**, **SubType**, and **Type** properties.


## Example

The following example adds a shape to the first slide of the active presentation and sets a filter effect animation behavior.


```vb
Sub ChangeFilterEffect()

    Dim sldFirst As Slide
    Dim shpHeart As Shape
    Dim effNew As Effect
    Dim bhvEffect As AnimationBehavior

    Set sldFirst = ActivePresentation.Slides(1)

    Set shpHeart = sldFirst.Shapes.AddShape(Type:=msoShapeHeart, _
        Left:=100, Top:=100, Width:=100, Height:=100)

    Set effNew = sldFirst.TimeLine.MainSequence.AddEffect _
        (Shape:=shpHeart, EffectID:=msoAnimEffectChangeFillColor, _
        Trigger:=msoAnimTriggerAfterPrevious)

    Set bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter)

    With bhvEffect.FilterEffect
         .Type = msoAnimFilterEffectTypeWipe
         .Subtype = msoAnimFilterEffectSubtypeUp
         .Reveal = msoTrue
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

