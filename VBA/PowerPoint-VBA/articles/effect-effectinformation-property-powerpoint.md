---
title: Effect.EffectInformation Property (PowerPoint)
keywords: vbapp10.chm652018
f1_keywords:
- vbapp10.chm652018
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.EffectInformation
ms.assetid: 68c61bfc-842e-6659-eda9-cc4899c50b94
ms.date: 06/08/2017
---


# Effect.EffectInformation Property (PowerPoint)

Returns an  **[EffectInformation](effectinformation-object-powerpoint.md)** object that represents information for a specified animation effect.


## Syntax

 _expression_. **EffectInformation**

 _expression_ A variable that represents an **Effect** object.


### Return Value

EffectInformation


## Example

This example adds a sound effect to the main animation sequence for a given shape.


```vb
Sub AddSoundEffect()

    Dim effMain As Effect

    Set effMain = ActivePresentation.Slides(1).TimeLine.MainSequence(1)

    MsgBox effMain.EffectInformation.AfterEffect

End Sub
```


## See also


#### Concepts


[Effect Object](effect-object-powerpoint.md)


