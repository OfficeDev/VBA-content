---
title: SlideShowTransition.Speed Property (PowerPoint)
keywords: vbapp10.chm539010
f1_keywords:
- vbapp10.chm539010
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.Speed
ms.assetid: 7c5b9dd2-88d3-5e34-619a-b35c3937a276
ms.date: 06/08/2017
---


# SlideShowTransition.Speed Property (PowerPoint)

Represents the speed of the transition to the specified slide. Read/write.


## Syntax

 _expression_. **Speed**

 _expression_ A variable that represents a **SlideShowTransition** object.


### Return Value

PpTransitionSpeed


## Remarks

The value of the  **Speed** property can be one of these **PpTransitionSpeed** constants.


||
|:-----|
|**ppTransitionSpeedFast**|
|**ppTransitionSpeedMedium**|
|**ppTransitionSpeedMixed**|
|**ppTransitionSpeedSlow**|

## Example

This example sets the special effect for the transition to the first slide in the active presentation and specifies that the transition be fast.


```vb
With ActivePresentation.Slides(1).SlideShowTransition

    .EntryEffect = ppEffectStripsDownLeft

    .Speed = ppTransitionSpeedFast

End With
```


## See also


#### Concepts


[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

