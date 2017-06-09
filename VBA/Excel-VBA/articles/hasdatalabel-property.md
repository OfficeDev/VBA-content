---
title: HasDataLabel Property
keywords: vbagr10.chm5207462
f1_keywords:
- vbagr10.chm5207462
ms.prod: excel
api_name:
- Excel.HasDataLabel
ms.assetid: d8fd8c48-4723-4da9-0b8a-82d02c93a19f
ms.date: 06/08/2017
---


# HasDataLabel Property

 **True** if the point has a data label. Read/write **Boolean**.


## Example

This example turns on the data label for point seven in series three, and then it sets the data label color to blue.


```vb
With myChart.SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels Type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```


