---
title: LeaderLines Object
keywords: vbagr10.chm5207590
f1_keywords:
- vbagr10.chm5207590
ms.prod: excel
api_name:
- Excel.LeaderLines
ms.assetid: 9704f195-dbbc-6979-c57d-8ced3557cdde
ms.date: 06/08/2017
---


# LeaderLines Object

Represents leader lines in the specified chart. Leader lines connect data labels to data points. This object isn't a collection; there's no object that represents a single leader line.


## Using the LeaderLines Object

Use the  **[LeaderLines](leaderlines-property.md)** property to return the  **LeaderLines** object. The following example adds data labels and blue leader lines to series one in the chart.


```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```


