---
title: Series.HasLeaderLines Property (Excel)
keywords: vbaxl10.chm578120
f1_keywords:
- vbaxl10.chm578120
ms.prod: excel
api_name:
- Excel.Series.HasLeaderLines
ms.assetid: 9401e5a6-5acc-7503-02e6-b6dc42f381bb
ms.date: 06/08/2017
---


# Series.HasLeaderLines Property (Excel)

 **True** if the series has leader lines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasLeaderLines**

 _expression_ A variable that represents a **Series** object.


## Remarks

This property applies only to pie charts.


## Example

This example adds data labels and blue leader lines to series one on the pie chart. If no leader lines are visible this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.


```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

