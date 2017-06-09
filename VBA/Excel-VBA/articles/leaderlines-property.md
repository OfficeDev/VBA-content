---
title: LeaderLines Property
keywords: vbagr10.chm5207592
f1_keywords:
- vbagr10.chm5207592
ms.prod: excel
api_name:
- Excel.LeaderLines
ms.assetid: ddd9ab86-d135-73de-b888-3ba43c39ece8
ms.date: 06/08/2017
---


# LeaderLines Property

Returns a  **LeaderLines** object that represents the leader lines for the specified series. Read-only.


## Example

This example adds data labels and blue leader lines to series one on the pie chart.


```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```


