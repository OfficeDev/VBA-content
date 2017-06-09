---
title: DataLabel Property
keywords: vbagr10.chm65694
f1_keywords:
- vbagr10.chm65694
ms.prod: excel
api_name:
- Excel.DataLabel
ms.assetid: 3f85b4c2-5b7d-311a-95f9-ad08d5c23e39
ms.date: 06/08/2017
---


# DataLabel Property

Returns a  **[DataLabel](datalabel-object.md)** object that represents the data label associated with the specified point or trendline. Read-only.


## Example

This example turns on the data label for point seven in series three, and then it sets the data label color to blue.


```vb
With myChart.SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```


