---
title: AnimationSettings.ChartUnitEffect Property (PowerPoint)
keywords: vbapp10.chm565016
f1_keywords:
- vbapp10.chm565016
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.ChartUnitEffect
ms.assetid: a2b66cf3-c8b9-6b9c-d184-13a828b474b2
ms.date: 06/08/2017
---


# AnimationSettings.ChartUnitEffect Property (PowerPoint)

Returns or sets a value that indicates whether the graph range is animated by series, category, or element. Read/write.


## Syntax

 _expression_. **ChartUnitEffect**

 _expression_ A variable that represents a **AnimationSettings** object.


### Return Value

PpChartUnitEffect


## Remarks

If your graph doesn't become animated, make sure that the  **[Animate](animationsettings-animate-property-powerpoint.md)** property is set to **True**

The value of the  **ChartUnitEffect** property can be one of these **PpChartUnitEffect** constants.


||
|:-----|
|**ppAnimateByCategory**|
|**ppAnimateByCategoryElements**|
|**ppAnimateBySeries**|
|**ppAnimateBySeriesElements**|
|**ppAnimateChartAllAtOnce**|
|**ppAnimateChartMixed**|

## Example

This example sets shape two on slide three in the active presentation to be animated by series. Shape two must be a graph for this to work.


```vb
With ActivePresentation.Slides(3).Shapes(2)

    With .AnimationSettings

        .ChartUnitEffect = ppAnimateBySeries

        .EntryEffect = ppEffectFlyFromLeft

        .Animate = True

    End With

End With
```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

