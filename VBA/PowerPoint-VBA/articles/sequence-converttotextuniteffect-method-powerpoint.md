---
title: Sequence.ConvertToTextUnitEffect Method (PowerPoint)
keywords: vbapp10.chm651012
f1_keywords:
- vbapp10.chm651012
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.ConvertToTextUnitEffect
ms.assetid: f6d2dabb-e8c5-99a9-5924-e897cbdc9968
ms.date: 06/08/2017
---


# Sequence.ConvertToTextUnitEffect Method (PowerPoint)

Returns an  **[Effect](effect-object-powerpoint.md)** object that represents how text should be animated.


## Syntax

 _expression_. **ConvertToTextUnitEffect**( **_Effect_**, **_unitEffect_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**|The animation effect to which the text unit effect applies.|
| _unitEffect_|Required|**[MsoAnimTextUnitEffect](msoanimtextuniteffect-enumeration-powerpoint.md)**|How the text should be animated.|

### Return Value

Effect


## Example

This example adds an animation to a given shape and animates its accompanying text by character.


```vb
Sub NewTextUnitEffect()

    Dim shpFirst As Shape
    Dim tmlMain As TimeLine

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)
    Set tmlMain = ActivePresentation.Slides(1).TimeLine

    tmlMain.MainSequence.ConvertToTextUnitEffect _
        Effect:=tmlMain.MainSequence.AddEffect(Shape:=shpFirst, _
            EffectID:=msoAnimEffectRandomEffects), _
        unitEffect:=msoAnimTextUnitEffectByCharacter

End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

