---
title: HasLegend Property
keywords: vbagr10.chm65589
f1_keywords:
- vbagr10.chm65589
ms.prod: excel
api_name:
- Excel.HasLegend
ms.assetid: b4dbef39-9d83-2f6e-fe06-8ca38cceeeec
ms.date: 06/08/2017
---


# HasLegend Property

 **True** if the chart has a legend. Read/write **Boolean**.


## Example

This example turns on the legend for the chart and then sets the legend font color to blue.


```vb
With myChart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
End With
```


