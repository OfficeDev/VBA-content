---
title: GapDepth Property
keywords: vbagr10.chm5207398
f1_keywords:
- vbagr10.chm5207398
ms.prod: excel
api_name:
- Excel.GapDepth
ms.assetid: 0aa59fe6-29bf-c014-8c11-18481f9c5603
ms.date: 06/08/2017
---


# GapDepth Property

Returns or sets the distance between the data series on a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write  **Long**.


## Example

This example sets the distance between the data series to 200 percent of the marker width. The example should be run on a 3-D chart (the  **GapDepth** property fails on 2-D charts).


```
myChart.GapDepth = 200
```


