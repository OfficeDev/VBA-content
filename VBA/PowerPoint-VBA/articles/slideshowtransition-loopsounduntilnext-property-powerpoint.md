---
title: SlideShowTransition.LoopSoundUntilNext Property (PowerPoint)
keywords: vbapp10.chm539008
f1_keywords:
- vbapp10.chm539008
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.LoopSoundUntilNext
ms.assetid: 64555d1a-20d2-cb4f-6168-dc9e9594e059
ms.date: 06/08/2017
---


# SlideShowTransition.LoopSoundUntilNext Property (PowerPoint)

Specifies whether the sound that's been set for the specified slide transition loops until the next sound starts. Read/write.


## Syntax

 _expression_. **LoopSoundUntilNext**

 _expression_ A variable that represents a **SlideShowTransition** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LoopSoundUntilNext** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The sound that's been set for the specified slide transition does not loop.|
|**msoTrue**| The sound that's been set for the specified slide transition loops until the next sound starts.|

## Example

This example specifies that the file Dudududu.wav will start to play at the transition to slide two in the active presentation and will continue to play until the next sound starts.


```vb
With ActivePresentation.Slides(2).SlideShowTransition

    .SoundEffect.ImportFromFile "c:\sndsys\dudududu.wav"

    .LoopSoundUntilNext = msoTrue

End With
```


## See also


#### Concepts


[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

