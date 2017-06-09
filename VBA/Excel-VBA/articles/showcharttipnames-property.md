---
title: ShowChartTipNames Property
keywords: vbagr10.chm5207986
f1_keywords:
- vbagr10.chm5207986
ms.prod: excel
api_name:
- Excel.ShowChartTipNames
ms.assetid: 0281bd54-2dbb-086f-23f7-ac507e19e519
ms.date: 06/08/2017
---


# ShowChartTipNames Property

 **True** if charts show chart tip names. The default value is **True**. Read/write  **Boolean**.


## Example

This example turns off chart tip names and values.


```vb
With myChart.Application 
 .ShowChartTipNames = False 
 .ShowChartTipValues = False 
End With
```


