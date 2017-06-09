---
title: FilterEffect.Subtype Property (PowerPoint)
keywords: vbapp10.chm669004
f1_keywords:
- vbapp10.chm669004
ms.prod: powerpoint
api_name:
- PowerPoint.FilterEffect.Subtype
ms.assetid: 1c244c97-9d50-93eb-7abc-5082aafcfb3e
ms.date: 06/08/2017
---


# FilterEffect.Subtype Property (PowerPoint)

 Sets or returns the subtype of the filter effect. Read/write.


## Syntax

 _expression_. **Subtype**

 _expression_ A variable that represents a **FilterEffect** object.


### Return Value

MsoAnimFilterEffectSubtype


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


[FilterEffect Object](filtereffect-object-powerpoint.md)

