---
title: NameIsAuto Property
keywords: vbagr10.chm65724
f1_keywords:
- vbagr10.chm65724
ms.prod: excel
api_name:
- Excel.NameIsAuto
ms.assetid: 92a06cde-f3fc-cc5b-9af9-0ec9545b90a8
ms.date: 06/08/2017
---


# NameIsAuto Property

 **True** if Microsoft Graph automatically determines the name of the trendline. Read/write **Boolean**.


## Example

This example sets Microsoft Graph to automatically determine the name for trendline one. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
myChart.SeriesCollection(1).Trendlines(1).NameIsAuto = True
```


