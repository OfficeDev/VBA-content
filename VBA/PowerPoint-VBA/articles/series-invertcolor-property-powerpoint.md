---
title: Series.InvertColor Property (PowerPoint)
keywords: vbapp10.chm716007
f1_keywords:
- vbapp10.chm716007
ms.prod: powerpoint
api_name:
- PowerPoint.Series.InvertColor
ms.assetid: e2ca8473-11d0-98fe-587e-740f7a00e85b
ms.date: 06/08/2017
---


# Series.InvertColor Property (PowerPoint)

Returns or sets the fill color for negative data points in a series. Read/write.


## Syntax

 _expression_. **InvertColor**

 _expression_ A variable that represents a **Series** object.


### Return Value

 **Integer**


## Remarks

The  **InvertColor** property enables you to set the fill color for negative data points as a specific numeric, hexadecimal, octal, or RGB color value. To set the value as an RBG value, use the Visual Basic RGB function. Instead of using the **InvertColor** property, you can use the **InvertColorIndex** property, which uses a simplier set of integer values from the current color palette.

For the  **InvertColor** property to have an effect, the **InvertIfNegative** property of the **Series** object must also be set to **True**.


## Example

The following code example sets the fill color of negative data points in the first series of chart 2 to magenta.


```vb
ActiveSheet.ChartObjects("Chart 2").Activate

ActiveChart.SeriesCollection(1).InvertIfNegative = True

ActiveChart.SeriesCollection(1).InvertColor = RGB(255, 0, 255)
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

