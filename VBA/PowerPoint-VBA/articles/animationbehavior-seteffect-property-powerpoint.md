---
title: AnimationBehavior.SetEffect Property (PowerPoint)
keywords: vbapp10.chm657015
f1_keywords:
- vbapp10.chm657015
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.SetEffect
ms.assetid: d23fe7c5-9b1b-f7c6-32d5-dd6fa00cb533
ms.date: 06/08/2017
---


# AnimationBehavior.SetEffect Property (PowerPoint)

Returns a  **SetEffect** object for the animation behavior. Read-only.


## Syntax

 _expression_. **SetEffect**

 _expression_ A variable that represents a **AnimationBehavior** object.


### Return Value

SetEffect


## Remarks

You can use the  **SetEffect** object returned to set the value of a property.


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


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

