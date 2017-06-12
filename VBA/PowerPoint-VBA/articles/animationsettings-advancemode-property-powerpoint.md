---
title: AnimationSettings.AdvanceMode Property (PowerPoint)
keywords: vbapp10.chm565008
f1_keywords:
- vbapp10.chm565008
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AdvanceMode
ms.assetid: 794d867f-cd7d-eeb6-0d6c-081e2be72ee5
ms.date: 06/08/2017
---


# AnimationSettings.AdvanceMode Property (PowerPoint)

Returns or sets a value that indicates whether the specified shape animation advances only when clicked or automatically after a specified amount of time. Read/write. 


## Syntax

 _expression_. **AdvanceMode**

 _expression_ A variable that represents an **AnimationSettings** object.


## Remarks

If your shape doesn't become animated, make sure that the  **[TextLevelEffect](animationsettings-textleveleffect-property-powerpoint.md)** property is set to a value other than **ppAnimateLevelNone** and that the **[Animate](animationsettings-animate-property-powerpoint.md)** property is set to **True**.

The value of the  **AdvanceMode** property can be one of these **PpAdvanceMode** constants.


||
|:-----|
|**ppAdvanceModeMixed**|
|**ppAdvanceOnClick**|
|**ppAdvanceOnTime**|

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

