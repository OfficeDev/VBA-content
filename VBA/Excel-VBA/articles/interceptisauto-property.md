---
title: InterceptIsAuto Property
keywords: vbagr10.chm65723
f1_keywords:
- vbagr10.chm65723
ms.prod: excel
api_name:
- Excel.InterceptIsAuto
ms.assetid: fd5b2155-8b45-8a67-19c9-8a18a4d3f6f3
ms.date: 06/08/2017
---


# InterceptIsAuto Property

 **True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean**.


## Remarks

Setting the  **[Intercept](intercept-property.md)** property sets this property to  **False**.


## Example

This example sets Microsoft Graph to automatically determine the trendline intercept point. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
myChart.SeriesCollection(1).Trendlines(1) _ 
 .InterceptIsAuto = True
```


