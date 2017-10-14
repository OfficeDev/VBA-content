---
title: HasDataLabels Property
keywords: vbagr10.chm65614
f1_keywords:
- vbagr10.chm65614
ms.prod: excel
api_name:
- Excel.HasDataLabels
ms.assetid: 1aa1d13e-69ec-0dab-1820-437c09afe820
ms.date: 06/08/2017
---


# HasDataLabels Property

 **True** if the series has data labels. Read/write **Boolean**.


## Example

This example turns on data labels for series three.


```vb
With myChart.SeriesCollection(3) 
 .HasDataLabels = True 
 .ApplyDataLabels Type:=xlValue 
End With
```


