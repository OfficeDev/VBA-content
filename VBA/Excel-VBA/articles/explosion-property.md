---
title: Explosion Property
ms.prod: excel
api_name:
- Excel.Explosion
ms.assetid: 252a3533-28df-4317-8af1-7509339409a5
ms.date: 06/08/2017
---


# Explosion Property

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write  **Long**.


## Example

This example sets the explosion value for point two. The example should be run on a pie chart.


```
myChart.SeriesCollection(1).Points(2). Explosion = 20

```


