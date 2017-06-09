---
title: AnimationBehavior.ColorEffect Property (PowerPoint)
keywords: vbapp10.chm657007
f1_keywords:
- vbapp10.chm657007
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.ColorEffect
ms.assetid: a1f8db9a-addf-c3f4-e5e3-0cc4b3f9f606
ms.date: 06/08/2017
---


# AnimationBehavior.ColorEffect Property (PowerPoint)

Returns a  **[ColorEffect](coloreffect-object-powerpoint.md)** object that represents the color properties for a specified animation behavior.


## Syntax

 _expression_. **ColorEffect**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

ColorEffect


## Example

This example adds a shape to the first slide of the active presentation and sets a color effect behavior to change the fill color of the new shape.


```vb
Sub ChangeColorEffect()

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

    Set bhvEffect = effNew.Behaviors.Add(Type:=msoAnimTypeColor)

    With bhvEffect.ColorEffect
        .From.RGB = RGB(Red:=255, Green:=0, Blue:=0)
        .To.RGB = RGB(Red:=0, Green:=0, Blue:=255)
    End With

End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

