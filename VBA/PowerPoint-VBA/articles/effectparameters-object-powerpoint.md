---
title: EffectParameters Object (PowerPoint)
keywords: vbapp10.chm654000
f1_keywords:
- vbapp10.chm654000
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters
ms.assetid: 78145783-800b-433b-25c2-54dd65f59556
ms.date: 06/08/2017
---


# EffectParameters Object (PowerPoint)

Represents various animation parameters for an  **[Effect](effect-object-powerpoint.md)** object, such as colors, fonts, sizes, and directions.


## Example

Use the [EffectParameters](effect-effectparameters-property-powerpoint.md)property of the  **Effect** object to return an **EffectParameters** object. The following example creates a shape, sets a fill effect, and changes the starting and ending fill colors.


```vb
Sub effParam()

    Dim shpNew As Shape
    Dim effNew As Effect

    Set shpNew = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeHeart, Left:=100, _
        Top:=100, Width:=150, Height:=150)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpNew, EffectID:=msoAnimEffectChangeFillColor, _
        Trigger:=msoAnimTriggerAfterPrevious)

    With effNew.EffectParameters
        .Color1.RGB = RGB(Red:=0, Green:=0, Blue:=255)
        .Color2.RGB = RGB(Red:=255, Green:=0, Blue:=0)
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


