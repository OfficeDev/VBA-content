---
title: AnimationSettings.AnimationOrder Property (PowerPoint)
keywords: vbapp10.chm565007
f1_keywords:
- vbapp10.chm565007
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AnimationOrder
ms.assetid: 0a29fb35-1cd8-4d12-184e-1132494a0864
ms.date: 06/08/2017
---


# AnimationSettings.AnimationOrder Property (PowerPoint)

Returns or sets an integer that represents the position of the specified shape within the collection of shapes to be animated. Read/write.


## Syntax

 _expression_. **AnimationOrder**

 _expression_ A variable that represents an **AnimationSettings** object.


### Return Value

Long


## Remarks

You won't see effects of setting this property unless the specified shape gets animated. For a shape to be animated, the  **TextLevelEffect** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone** and the **[Animate](animationsettings-animate-property-powerpoint.md)** property must be set to **True**.


 **Note**  Setting the  **AnimationOrder** property to a value that is less than the greatest existing **AnimationOrder** property value can shift the animation order.


## Example

This example specifies that shape two on slide two in the active presentation be animated second.


```vb
ActivePresentation.Slides(2).Shapes(2) _
    .AnimationSettings.AnimationOrder = 2
```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

