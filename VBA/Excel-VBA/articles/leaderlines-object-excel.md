---
title: LeaderLines Object (Excel)
keywords: vbaxl10.chm605072
f1_keywords:
- vbaxl10.chm605072
ms.prod: excel
api_name:
- Excel.LeaderLines
ms.assetid: ff4954f1-6967-9dd8-c9d6-8927d079e995
ms.date: 06/08/2017
---


# LeaderLines Object (Excel)

Represents leader lines on a chart. Leader lines connect data labels to data points.


## Remarks

 This object isn't a collection; there's no object that represents a single leader line.

This object applies only to pie charts.


## Example

Use the  **[LeaderLines](series-leaderlines-property-excel.md)** property to return the **LeaderLines** object. The following example adds data labels and blue leader lines to series one on chart one. If no leader lines are visible this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.


```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

