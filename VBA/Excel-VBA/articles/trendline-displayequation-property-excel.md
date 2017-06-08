---
title: Trendline.DisplayEquation Property (Excel)
keywords: vbaxl10.chm594079
f1_keywords:
- vbaxl10.chm594079
ms.prod: excel
api_name:
- Excel.Trendline.DisplayEquation
ms.assetid: a9c3de54-5690-bf9b-505a-65b069195d53
ms.date: 06/08/2017
---


# Trendline.DisplayEquation Property (Excel)

 **True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to **True** automatically turns on data labels. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayEquation**

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

