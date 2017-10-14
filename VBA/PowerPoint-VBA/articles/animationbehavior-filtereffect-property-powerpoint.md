---
title: AnimationBehavior.FilterEffect Property (PowerPoint)
keywords: vbapp10.chm657014
f1_keywords:
- vbapp10.chm657014
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.FilterEffect
ms.assetid: e661aea9-f83d-db2e-6988-4bc1f9e15287
ms.date: 06/08/2017
---


# AnimationBehavior.FilterEffect Property (PowerPoint)

Returns a  **FilterEffect** object that represents a filter effect for an animation behavior. Read-only.


## Syntax

 _expression_. **FilterEffect**

 _expression_ A variable that represents a **AnimationBehavior** object.


### Return Value

FilterEffect


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


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

