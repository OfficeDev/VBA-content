---
title: Effect.EffectParameters Property (PowerPoint)
keywords: vbapp10.chm652011
f1_keywords:
- vbapp10.chm652011
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.EffectParameters
ms.assetid: 18f43203-a16e-7779-923c-7da076d2943e
ms.date: 06/08/2017
---


# Effect.EffectParameters Property (PowerPoint)

Returns an  **[EffectParameters](effectparameters-object-powerpoint.md)** object that represents animation effect properties.


## Syntax

 _expression_. **EffectParameters**

 _expression_ A variable that represents an **Effect** object.


### Return Value

EffectParameters


## Example

This example adds an effect to the main animation sequence on the first slide. This effect changes the font for the first shape to Broadway.


```vb
Sub ChangeFontName()

    Dim shpText As Shape
    Dim effNew As Effect

    Set shpText = ActivePresentation.Slides(1).Shapes(1)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpText, EffectId:=msoAnimEffectChangeFont)

    effNew.EffectParameters.FontName = "Broadway"

End Sub
```


## See also


#### Concepts


[Effect Object](effect-object-powerpoint.md)


