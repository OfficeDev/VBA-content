---
title: AnimationBehavior.MotionEffect Property (PowerPoint)
keywords: vbapp10.chm657006
f1_keywords:
- vbapp10.chm657006
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.MotionEffect
ms.assetid: ef9601ab-7a01-ba03-a5ef-a50c4d2c3c79
ms.date: 06/08/2017
---


# AnimationBehavior.MotionEffect Property (PowerPoint)

Returns a  **[MotionEffect](motioneffect-object-powerpoint.md)** object that represents the properties of a motion animation.


## Syntax

 _expression_. **MotionEffect**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

MotionEffect


## Example

This example adds a new motion behavior to the first slide's main sequence that moves the specified animation sequence from one side of the page to the shape's original position.


```vb
Sub NewMotion()

    With ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
            .Behaviors.Add(msoAnimTypeMotion).MotionEffect
        .FromX = 100
        .FromY = 100
        .ToX = 0
        .ToY = 0
    End With

End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

