---
title: AnimationBehavior.Additive Property (PowerPoint)
keywords: vbapp10.chm657003
f1_keywords:
- vbapp10.chm657003
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.Additive
ms.assetid: 29dabc4f-a333-9b11-97a5-36237a95dcb0
ms.date: 06/08/2017
---


# AnimationBehavior.Additive Property (PowerPoint)

Sets or returns whether the current animation behavior is combined with other running animations. Read/write.


## Syntax

 _expression_. **Additive**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

MsoAnimAdditive


## Remarks

The value of the  **Additive** property can be one of these **MsoAnimAdditive** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoAnimAdditiveAddBase**|Does not combine current animation with other animations. The default.|
|**msoAnimAdditiveAddSum**| Combines the current animation with other running animations.|
Combining animation behaviors is particularly useful for rotation effects. For example, if the current animation changes rotation and another animation is also changing rotation, if this property is set to  **msoAnimAdditiveAddSum**, Microsoft PowerPoint adds together the rotations from both the animations.


## Example

The following example allows the current animation behavior to be added to another animation behavior.


```vb
Sub SetAdditive()

    Dim animBehavior As AnimationBehavior

    Set animBehavior = ActiveWindow.Selection.SlideRange(1) _
        .TimeLine.MainSequence(1).Behaviors(1)

    animBehavior.Additive = msoAnimAdditiveAddSum

End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

