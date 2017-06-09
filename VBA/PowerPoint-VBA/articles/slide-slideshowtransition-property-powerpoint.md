---
title: Slide.SlideShowTransition Property (PowerPoint)
keywords: vbapp10.chm531005
f1_keywords:
- vbapp10.chm531005
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.SlideShowTransition
ms.assetid: bb931628-0ad1-e58b-9ddb-5680cb6ce9ec
ms.date: 06/08/2017
---


# Slide.SlideShowTransition Property (PowerPoint)

Returns a  **[SlideShowTransition](slideshowtransition-object-powerpoint.md)** object that represents the special effects for the specified slide transition. Read-only.


## Syntax

 _expression_. **SlideShowTransition**

 _expression_ A variable that represents a **Slide** object.


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


[Slide Object](slide-object-powerpoint.md)

