---
title: EffectInformation.SoundEffect Property (PowerPoint)
keywords: vbapp10.chm655009
f1_keywords:
- vbapp10.chm655009
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.SoundEffect
ms.assetid: ff881716-307e-4cce-7cc5-68d32350527f
ms.date: 06/08/2017
---


# EffectInformation.SoundEffect Property (PowerPoint)

Returns a  **SoundEffect** object that represents the sound to be played during the transition to the specified slide. Read-only.


## Syntax

 _expression_. **SoundEffect**

 _expression_ A variable that represents an **EffectInformation** object.


### Return Value

SoundEffect


## Example

This example sets the file Bass.wav to be played whenever shape one on slide one in the active presentation is animated.


```vb
With ActivePresentation.Slides(1).Shapes(1).AnimationSettings

    .Animate = True

    .TextLevelEffect = ppAnimateByAllLevels

    .SoundEffect.ImportFromFile "c:\bass.wav"

End With
```


## See also


#### Concepts


[EffectInformation Object](effectinformation-object-powerpoint.md)


