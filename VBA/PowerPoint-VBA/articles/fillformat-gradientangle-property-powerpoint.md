---
title: FillFormat.GradientAngle Property (PowerPoint)
keywords: vbapp10.chm552034
f1_keywords:
- vbapp10.chm552034
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientAngle
ms.assetid: eb5362f0-5d3b-0091-7a83-0a8d58d90438
ms.date: 06/08/2017
---


# FillFormat.GradientAngle Property (PowerPoint)

Returns or sets the angle of the gradient fill for the specified fill format. Read/write.


## Syntax

 _expression_. **GradientAngle**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

 **Single**


## Remarks

A gradient fill can be specified in the formatting for various elements (shapes) in a chart. For example, you can use the  **Format Data Series** dialog box to format the columns in a **Column** chart to a gradient fill. In this case, the **GradientAngle** property corresponds to the setting of the **Angle** box in the **Fill** category of the **Format Data Series** dialog box. The valid range of values for the **GradientAngle** property is from 0 through 359.9.


## Example

The following example sets the angle of the gradient fill for series 1 in chart 1 to 45 degrees.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate

ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

