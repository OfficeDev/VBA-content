---
title: Intercept Property
keywords: vbagr10.chm65722
f1_keywords:
- vbagr10.chm65722
ms.prod: excel
api_name:
- Excel.Intercept
ms.assetid: 9c7c4193-8f9d-0f33-74c7-055a9124320e
ms.date: 06/08/2017
---


# Intercept Property

Returns or sets the point where the trendline crosses the value axis. Read/write  **Double**.


## Remarks

Setting this property sets the  **[InterceptIsAuto](interceptisauto-property.md)** property to  **False**.


## Example

This example sets trendline one to cross the value axis at 5. The example should be run on a 2-D column chart that contains a single series with a trendline.


```
myChart.SeriesCollection(1).Trendlines(1).Intercept = 5
```


