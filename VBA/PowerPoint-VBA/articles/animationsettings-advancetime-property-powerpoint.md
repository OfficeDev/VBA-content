---
title: AnimationSettings.AdvanceTime Property (PowerPoint)
keywords: vbapp10.chm565009
f1_keywords:
- vbapp10.chm565009
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AdvanceTime
ms.assetid: f4e5cec6-ba11-f605-3b3f-c4867fbce315
ms.date: 06/08/2017
---


# AnimationSettings.AdvanceTime Property (PowerPoint)

Returns or sets the amount of time, in seconds, after which the specified shape will become animated. Read/write.


## Syntax

 _expression_. **AdvanceTime**

 _expression_ A variable that represents an **AnimationSettings** object.


### Return Value

Single


## Remarks

The specified slide animation won't start automatically after the amount of time you've specified unless the  **[AdvanceMode](slideshowsettings-advancemode-property-powerpoint.md)** property of the animation is set to **ppAdvanceOnTime**.


## Example

This example sets shape two on slide one in the active presentation to become animated automatically after five seconds.


```vb
With ActivePresentation.Slides(1).Shapes(2).AnimationSettings

    .AdvanceMode = ppAdvanceOnTime

    .AdvanceTime = 5

    .TextLevelEffect = ppAnimateByAllLevels

    .Animate = True

End With


```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

