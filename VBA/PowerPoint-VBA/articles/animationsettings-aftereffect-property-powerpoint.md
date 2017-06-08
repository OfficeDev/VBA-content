---
title: AnimationSettings.AfterEffect Property (PowerPoint)
keywords: vbapp10.chm565006
f1_keywords:
- vbapp10.chm565006
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AfterEffect
ms.assetid: d8ccab29-8637-a48d-0f44-81a7fd1cca0b
ms.date: 06/08/2017
---


# AnimationSettings.AfterEffect Property (PowerPoint)

Returns or sets a  **PpAfterEffect** constant that indicates whether the specified shape appears dimmed, hidden, or unchanged after it is built. Read/write.


## Syntax

 _expression_. **AfterEffect**

 _expression_ A variable that represents an **AnimationSettings** object.


## Remarks

You won't see the aftereffect you set for a shape unless the shape gets animated and at least one other shape on the slide gets animated after it. For a shape to be animated, the  **[TextLevelEffect](animationsettings-textleveleffect-property-powerpoint.md)** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone**, or the **[EntryEffect](animationsettings-entryeffect-property-powerpoint.md)** property must be set to a constant other than **ppEffectNone**. In addition, the **[Animate](animationsettings-animate-property-powerpoint.md)** property must be set to **True**. To change the build order of the shapes on a slide, use the **[AnimationOrder](animationsettings-animationorder-property-powerpoint.md)** property.

The value of the  **AfterEffect** property can be one of these **PpAfterEffect** constants.


||
|:-----|
|**ppAfterEffectDim**|
|**ppAnimAfterEffectHide**|
|**ppAfterEffectHideOnNextClick**|
|**ppAfterEffectMixed**|
|**ppAfterEffectNone**|

## Example

This example specifies that the title on slide one in the active presentation is to appear dimmed after the title is built. If the title is the last or only shape to be built on slide one, the text does not appear dimmed.


```vb
With ActivePresentation.Slides(1).Shapes.Title.AnimationSettings

    .Animate = True

    .TextLevelEffect = ppAnimateByAllLevels

    .AfterEffect = ppAfterEffectDim

End With
```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

