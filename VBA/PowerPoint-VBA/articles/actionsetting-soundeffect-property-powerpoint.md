---
title: ActionSetting.SoundEffect Property (PowerPoint)
keywords: vbapp10.chm567009
f1_keywords:
- vbapp10.chm567009
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.SoundEffect
ms.assetid: ea577e7a-32be-ec68-42ab-625816534ab4
ms.date: 06/08/2017
---


# ActionSetting.SoundEffect Property (PowerPoint)

Returns a  **SoundEffect** object that represents the sound to be played during the transition to the specified slide. Read-only.


## Syntax

 _expression_. **SoundEffect**

 _expression_ A variable that represents an **ActionSetting** object.


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


[ActionSetting Object](actionsetting-object-powerpoint.md)

