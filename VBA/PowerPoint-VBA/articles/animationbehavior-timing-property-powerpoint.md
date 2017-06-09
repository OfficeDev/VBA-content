---
title: AnimationBehavior.Timing Property (PowerPoint)
keywords: vbapp10.chm657011
f1_keywords:
- vbapp10.chm657011
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.Timing
ms.assetid: 343f11d4-04bf-2637-dbbc-dc3256d57940
ms.date: 06/08/2017
---


# AnimationBehavior.Timing Property (PowerPoint)

Returns a  **[Timing](timing-object-powerpoint.md)** object that represents the timing properties for an animation sequence.


## Syntax

 _expression_. **Timing**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

Timing


## Example

The following example sets the duration of the first animation sequence on the first slide.


```vb
Sub SetTiming()
    ActivePresentation.Slides(1).TimeLine _
        .MainSequence(1).Timing.Duration = 1
End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

