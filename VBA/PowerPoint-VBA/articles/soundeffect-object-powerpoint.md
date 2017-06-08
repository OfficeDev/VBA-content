---
title: SoundEffect Object (PowerPoint)
keywords: vbapp10.chm540000
f1_keywords:
- vbapp10.chm540000
ms.prod: powerpoint
api_name:
- PowerPoint.SoundEffect
ms.assetid: 216e8bed-e6d7-e751-4d53-1c9902ddb89f
ms.date: 06/08/2017
---


# SoundEffect Object (PowerPoint)

Represents the sound effect that accompanies an animation or slide transition in a slide show.


## Example

Use the [SoundEffect](animationsettings-soundeffect-property-powerpoint.md)property of the  **AnimationSettings** object to return the **SoundEffect** object that represents the sound effect that accompanies an animation. The following example specifies that the animation of the title on slide one in the active presentation be accompanied by the sound in the Bass.wav file.


```vb
With ActivePresentation.Slides(1).Shapes(1).AnimationSettings

    .TextLevelEffect = ppAnimateByAllLevels

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

End With
```

Use the  **SoundEffect** property of the **SlideShowTransition** object to return the **SoundEffect** object that represents the sound effect that accompanies a slide transition.

The following example specifies that the transition to slide one in the active presentation be accompanied by the sound in the Bass.wav file.




```vb
ActivePresentation.Slides(1).SlideShowTransition.SoundEffect _
    .ImportFromFile "c:\sndsys\bass.wav"
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

