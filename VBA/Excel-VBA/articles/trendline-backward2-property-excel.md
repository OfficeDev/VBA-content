---
title: Trendline.Backward2 Property (Excel)
keywords: vbaxl10.chm594091
f1_keywords:
- vbaxl10.chm594091
ms.prod: excel
api_name:
- Excel.Trendline.Backward2
ms.assetid: 28712c4d-7772-d61e-0151-22eea8ff6383
ms.date: 06/08/2017
---


# Trendline.Backward2 Property (Excel)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write  **Double** .


## Syntax

 _expression_ . **Backward2**

 _expression_ A variable that represents a **Trendline** object.


## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With Charts("Chart1").SeriesCollection(1).Trendlines(1) 
 .Forward2 = 5 
 .Backward2 = .5 
End With 

```


## See also


#### Concepts


[Trendline Object](trendline-object-excel.md)

