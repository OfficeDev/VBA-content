---
title: Series.InvertColorIndex Property (PowerPoint)
keywords: vbapp10.chm716008
f1_keywords:
- vbapp10.chm716008
ms.prod: powerpoint
api_name:
- PowerPoint.Series.InvertColorIndex
ms.assetid: 879637a8-52a7-a6ac-a882-386dad1808cb
ms.date: 06/08/2017
---


# Series.InvertColorIndex Property (PowerPoint)

Returns or sets the fill color for negative data points in a series. Read/write.


## Syntax

 _expression_. **InvertColorIndex**

 _expression_ A variable that represents a **Series** object.


### Return Value

 **Integer**


## Remarks

The  **InvertColorIndex** property enables you to set the fill color for negative data points as a color index value from 0 to 56. Instead of using the **InvertColorIndex** property, you can use the **InvertColor** property, which enables you to set the color as a specific numeric, hexadecimal, octal, or RGB color value.

For the  **InvertColorIndex** property to have an effect, the **InvertIfNegative** property of the **Series** object must also be set to **True**.


## Example

The following code example sets the fill color of negative data points in the first series of chart 2 to magenta.


```vb
ActiveSheet.ChartObjects("Chart 2").Activate

ActiveChart.SeriesCollection(1).InvertIfNegative = True

ActiveChart.SeriesCollection(1).InvertColorIndex = 7
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

