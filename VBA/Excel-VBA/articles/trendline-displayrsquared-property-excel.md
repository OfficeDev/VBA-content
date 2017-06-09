---
title: Trendline.DisplayRSquared Property (Excel)
keywords: vbaxl10.chm594080
f1_keywords:
- vbaxl10.chm594080
ms.prod: excel
api_name:
- Excel.Trendline.DisplayRSquared
ms.assetid: e8e447c3-d379-f6d0-74f2-629fa53b42ef
ms.date: 06/08/2017
---


# Trendline.DisplayRSquared Property (Excel)

 **True** if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to **True** automatically turns on data labels. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRSquared**

 _expression_ A variable that represents a **Trendline** object.


## Example

This example displays the R-squared value and equation for trendline one in Chart1. The example should be run on a 2-D column chart that has a trendline for the first series.


```vb
With Charts("Chart1").SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-excel.md)

