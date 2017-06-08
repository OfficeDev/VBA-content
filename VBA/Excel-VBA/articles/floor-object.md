---
title: Floor Object
keywords: vbagr10.chm5207374
f1_keywords:
- vbagr10.chm5207374
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: ce76e68b-7b15-7e2c-4464-07befbf53cc5
ms.date: 06/08/2017
---


# Floor Object

Represents the floor of the specified 3-D chart.


## Using the Floor Object

Use the  **Floor** property to return the **Floor** object. The following example sets the floor color for the chart to cyan. If the chart isn't a 3-D chart, this example will fail.


```
myChart.Floor.Interior.Color = RGB(0, 255, 255)
```


