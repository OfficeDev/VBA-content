---
title: SlideShowTransition.SoundEffect Property (PowerPoint)
keywords: vbapp10.chm539009
f1_keywords:
- vbapp10.chm539009
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.SoundEffect
ms.assetid: 69cff9a7-777a-57a0-d897-f132ba028bdd
ms.date: 06/08/2017
---


# SlideShowTransition.SoundEffect Property (PowerPoint)

Returns a  **SoundEffect** object that represents the sound to be played during the transition to the specified slide. Read-only.


## Syntax

 _expression_. **SoundEffect**

 _expression_ A variable that represents a **SlideShowTransition** object.


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


[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

