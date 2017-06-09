---
title: Master.SlideShowTransition Property (PowerPoint)
keywords: vbapp10.chm533016
f1_keywords:
- vbapp10.chm533016
ms.prod: powerpoint
api_name:
- PowerPoint.Master.SlideShowTransition
ms.assetid: 935cadd9-a57a-a792-62b4-e198546438b2
ms.date: 06/08/2017
---


# Master.SlideShowTransition Property (PowerPoint)

Returns a  **[SlideShowTransition](slideshowtransition-object-powerpoint.md)** object that represents the special effects for the specified slide transition. Read-only.


## Syntax

 _expression_. **SlideShowTransition**

 _expression_ A variable that represents a **Master** object.


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


[Master Object](master-object-powerpoint.md)

