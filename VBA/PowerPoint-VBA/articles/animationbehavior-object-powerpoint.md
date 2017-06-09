---
title: AnimationBehavior Object (PowerPoint)
keywords: vbapp10.chm657000
f1_keywords:
- vbapp10.chm657000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior
ms.assetid: 70eeb4aa-b9ba-ff7d-93ee-425cf191a6cb
ms.date: 06/08/2017
---


# AnimationBehavior Object (PowerPoint)

Represents the behavior of an animation effect, the main animation sequence, or an interactive animation sequence. The  **AnimationBehavior** object is a member of the **[AnimationBehaviors](animationbehaviors-object-powerpoint.md)** collection.


## Example

Use [Behaviors](effect-behaviors-property-powerpoint.md)(index), where index is the number of the behavior in the sequence of behaviors, to return a single  **AnimationBehavior** object. The following example sets the positions of the a rotation's starting and ending points. This example assumes that the first behavior for the main animation sequence is a **[RotationEffect](rotationeffect-object-powerpoint.md)** object.


```vb
Sub Change()
    With ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
            .Behaviors(1).RotationEffect
        .From = 1
        .To = 180
    End With
End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

