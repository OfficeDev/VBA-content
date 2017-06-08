---
title: ColorEffect Object (PowerPoint)
keywords: vbapp10.chm659000
f1_keywords:
- vbapp10.chm659000
ms.prod: powerpoint
api_name:
- PowerPoint.ColorEffect
ms.assetid: c21ca9cd-3aaa-35f7-3d0e-89ca155322f2
ms.date: 06/08/2017
---


# ColorEffect Object (PowerPoint)

Represents a color effect for an animation behavior.


## Example

Use the [ColorEffect](animationbehavior-coloreffect-property-powerpoint.md)property of the  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object to return a **ColorEffect** object. Color effects can be changed using the **ColorEffect** object's[From](coloreffect-from-property-powerpoint.md)and [To](coloreffect-to-property-powerpoint.md)properties, as shown below. Color effects are initially set using the  **To** property, and then can be changed by a specific number using the[By](coloreffect-by-property-powerpoint.md)property. The following example adds a shape to the first slide of the active presentation and sets a color effect animation behavior to change the fill color of the new shape.


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


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

