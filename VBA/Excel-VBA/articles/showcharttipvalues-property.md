---
title: ShowChartTipValues Property
keywords: vbagr10.chm5207990
f1_keywords:
- vbagr10.chm5207990
ms.prod: excel
api_name:
- Excel.ShowChartTipValues
ms.assetid: aeff428a-01c2-51cc-2397-e178386e3e69
ms.date: 06/08/2017
---


# ShowChartTipValues Property

 **True** if charts show chart tip values. The default value is **True**. Read/write  **Boolean**.


## Example

This example turns off chart tip names and values.


```vb
With myChart.Application 
 .ShowChartTipNames = False 
 .ShowChartTipValues = False 
End With
```


