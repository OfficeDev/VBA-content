---
title: AnimationPoints.Smooth Property (PowerPoint)
keywords: vbapp10.chm663005
f1_keywords:
- vbapp10.chm663005
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoints.Smooth
ms.assetid: cf41b527-91cc-81ac-ebb8-8fdf40bee5df
ms.date: 06/08/2017
---


# AnimationPoints.Smooth Property (PowerPoint)

Determines whether the transition from one animation point to another is smoothed. Read/write.


## Syntax

 _expression_. **Smooth**

 _expression_ A variable that represents a **AnimationPoints** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Smooth** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The animation point should not be smoothed.|
|**msoTrue**| The default. The animation should be smoothed.|

## Example

This example changes smoothing for an animation point.


```vb
Sub ChangeSmooth(ByVal ani As AnimationBehavior, ByVal bln As MsoTriState)

    ani.PropertyEffect.Points.Smooth = bln

End Sub
```


## See also


#### Concepts


[AnimationPoints Object](animationpoints-object-powerpoint.md)
[LegendKey Object](legendkey-object-powerpoint.md)
[Series Object](series-object-powerpoint.md)

