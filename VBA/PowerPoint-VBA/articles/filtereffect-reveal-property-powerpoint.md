---
title: FilterEffect.Reveal Property (PowerPoint)
keywords: vbapp10.chm669005
f1_keywords:
- vbapp10.chm669005
ms.prod: powerpoint
api_name:
- PowerPoint.FilterEffect.Reveal
ms.assetid: 01aaa4e5-e433-3e19-3f78-d266a1bf2890
ms.date: 06/08/2017
---


# FilterEffect.Reveal Property (PowerPoint)

Determines how the embedded objects will be revealed. Read/write.


## Syntax

 _expression_. **Reveal**

 _expression_ A variable that represents a **FilterEffect** object.


### Return Value

MsoTriState


## Remarks

Setting a value of  **msoTrue** for the **Reveal** property when the filter effect type is **msoAnimFilterEffectTypeWipe** will make the shape appear. Setting a value of **msoFalse** will make the object disappear. In other words, if your filter is set to wipe and **Reveal** is true, you will get a wipe-in effect, and when **Reveal** is false, you will get a wipe-out effect.

The value of the  **Reveal** property can be one of these **MsoTriState** constants.


||
|:-----|
|**msoFalse**|
|**msoTrue**|

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

