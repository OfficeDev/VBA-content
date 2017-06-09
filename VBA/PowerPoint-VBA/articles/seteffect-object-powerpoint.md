---
title: SetEffect Object (PowerPoint)
keywords: vbapp10.chm670000
f1_keywords:
- vbapp10.chm670000
ms.prod: powerpoint
api_name:
- PowerPoint.SetEffect
ms.assetid: 299eff64-54d6-3689-a031-ca6a3756afca
ms.date: 06/08/2017
---


# SetEffect Object (PowerPoint)

Represents a set effect for an animation behavior. You can use the  **SetEffect** object to set the value of a property.


## Remarks

Use the  **SetEffect** property of the **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object to return a **SetEffect** object. Set effects can be changed using the **SetEffect** object's **Property** and **To** properties.


## Example

The following example adds a shape to the first slide of the active presentation and sets a set effect animation behavior.


```vb
Sub ChangeSetEffect()

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

    Set bhvEffect = effNew.Behaviors.Add(msoAnimTypeSet)

    With bhvEffect.SetEffect
         .Property = msoAnimShapeFillColor
         .To = RGB(Red:=0, Green:=255, Blue:=255)
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

