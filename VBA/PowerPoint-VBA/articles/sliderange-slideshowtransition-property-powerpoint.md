---
title: SlideRange.SlideShowTransition Property (PowerPoint)
keywords: vbapp10.chm532005
f1_keywords:
- vbapp10.chm532005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.SlideShowTransition
ms.assetid: d97522ce-75c8-16f7-cdee-337b3af035db
ms.date: 06/08/2017
---


# SlideRange.SlideShowTransition Property (PowerPoint)

Returns a  **[SlideShowTransition](slideshowtransition-object-powerpoint.md)** object that represents the special effects for the specified slide transition. Read-only.


## Syntax

 _expression_. **SlideShowTransition**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

SlideShowTransition


## Example

This example sets slide two in the active presentation to advance automatically after 5 seconds during a slide show and to play a dog bark sound at the slide transition.


```vb
With ActivePresentation.Slides(2).SlideShowTransition

    .AdvanceOnTime = True

    .AdvanceTime = 5

    .SoundEffect.ImportFromFile "c:\windows\media\dogbark.wav"

End With

ActivePresentation.SlideShowSettings.AdvanceMode = _
    ppSlideShowUseSlideTimings
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

