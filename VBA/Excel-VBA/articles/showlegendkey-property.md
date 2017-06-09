---
title: ShowLegendKey Property
keywords: vbagr10.chm65707
f1_keywords:
- vbagr10.chm65707
ms.prod: excel
api_name:
- Excel.ShowLegendKey
ms.assetid: 508fe969-30fc-f313-2406-213b5d8594ff
ms.date: 06/08/2017
---


# ShowLegendKey Property

 **True** if the data label legend key is visible. Read/write **Boolean**.


## Example

This example sets the data labels for series one to show values and the legend key.


```vb
With myChart.SeriesCollection(1).DataLabels 
 .ShowLegendKey = True 
 .Type = xlShowValue 
End With
```


