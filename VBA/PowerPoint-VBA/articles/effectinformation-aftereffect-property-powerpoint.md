---
title: EffectInformation.AfterEffect Property (PowerPoint)
keywords: vbapp10.chm655003
f1_keywords:
- vbapp10.chm655003
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.AfterEffect
ms.assetid: 18fd4307-c737-2a97-09bc-ff381a18d768
ms.date: 06/08/2017
---


# EffectInformation.AfterEffect Property (PowerPoint)

Returns an  **PpAfterEffect** constant that indicates whether an after effect appears dimmed, hidden, or unchanged after it runs. Read-only.


## Syntax

 _expression_. **AfterEffect**

 _expression_ A variable that represents an **EffectInformation** object.


## Remarks

The value returned by the  **AfterEffect** property can be one of these **PpAfterEffect** constants.


||
|:-----|
|**ppAfterEffectDim**|
|**ppAnimAfterEffectHide**|
|**ppAfterEffectHideOnNextClick**|
|**ppAfterEffectMixed**|
|**ppAfterEffectNone**|

## Example

This example specifies that the title on slide one in the active presentation is to appear dimmed after the title is built. If the title is the last or only shape to be built on slide one, the text won't be dimmed.


```vb
With ActivePresentation.Slides(1).Shapes.Title.AnimationSettings

    .Animate = True

    .TextLevelEffect = ppAnimateByAllLevels

    .AfterEffect = ppAfterEffectDim

End With
```


## See also


#### Concepts



[EffectInformation Object](effectinformation-object-powerpoint.md)

