---
title: ShowSeriesName Property
keywords: vbagr10.chm67466
f1_keywords:
- vbagr10.chm67466
ms.prod: excel
api_name:
- Excel.ShowSeriesName
ms.assetid: 73374913-f0b9-501c-7516-4497d6b85977
ms.date: 06/08/2017
---


# ShowSeriesName Property

Allows the user to show the series name for the data labels on a chart. Read/write Boolean.

 _expression_. **ShowSeriesName**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the series name to be shown for the data labels of the first series on the first chart.


```vb
Sub UseSeriesName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowSeriesName = True 
 
End Sub
```


