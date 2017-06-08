---
title: SoundEffect.Play Method (PowerPoint)
keywords: vbapp10.chm540006
f1_keywords:
- vbapp10.chm540006
ms.prod: powerpoint
api_name:
- PowerPoint.SoundEffect.Play
ms.assetid: d0f598cb-2c3c-936b-42a2-326ead1e995b
ms.date: 06/08/2017
---


# SoundEffect.Play Method (PowerPoint)

Plays the specified sound effect.


## Syntax

 _expression_. **Play**

 _expression_ A variable that represents a **SoundEffect** object.


## Example

This example plays the sound effect that's been set for the transition to slide two in the active presentation.


```vb
ActivePresentation.Slides(2).SlideShowTransition.SoundEffect.Play
```


## See also


#### Concepts


[SoundEffect Object](soundeffect-object-powerpoint.md)
[Player Object](player-object-powerpoint.md)

