---
title: Sequence.AddEffect Method (PowerPoint)
keywords: vbapp10.chm651004
f1_keywords:
- vbapp10.chm651004
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.AddEffect
ms.assetid: fea5ac1e-83ae-2241-bf3a-8cfdd8354791
ms.date: 06/08/2017
---


# Sequence.AddEffect Method (PowerPoint)

Returns an  **[Effect](effect-object-powerpoint.md)** object that represents a new animation effect added to a sequence of animation effects.


## Syntax

 _expression_. **AddEffect**( **_Shape_**, **_effectId_**, **_Level_**, **_trigger_**, **_Index_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required|**[Shape](shape-object-powerpoint.md)**|The shape to which the animation effect is added.|
| _effectId_|Required|**[MsoAnimEffect](msoanimeffect-enumeration-powerpoint.md)**|The animation effect to be applied.|
| _Level_|Optional|**[MsoAnimateByLevel](msoanimatebylevel-enumeration-powerpoint.md)**|For charts, diagrams, or text, the level to which the animation effect will be applied. The default value is  **msoAnimationLevelNone**.|
| _trigger_|Optional|**[MsoAnimTriggerType](msoanimatebylevel-enumeration-powerpoint.md)**|The action that triggers the animation effect. The default value is  **msoAnimTriggerOnPageClick**.|
| _Index_|Optional|**Long**|The position at which the effect will be placed in the collection of animation effects. The default value is -1 (added to the end). |

### Return Value

Effect


## Example

The following example adds a bouncing animation to the first shape range on the first slide. This example assumes a shape range containing one or more shapes is selected on the first slide.


```vb
Sub AddBouncingAnimation()

    Dim sldActive As Slide
    Dim shpSelected As Shape

    Set sldActive = ActiveWindow.Selection.SlideRange(1)
    Set shpSelected = ActiveWindow.Selection.ShapeRange(1)

    ' Add a bouncing animation.
    sldActive.TimeLine.MainSequence.AddEffect _
        Shape:=shpSelected, effectId:=msoAnimEffectBounce

End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

