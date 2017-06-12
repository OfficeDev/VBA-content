---
title: ShowPercentage Property
keywords: vbagr10.chm3077090
f1_keywords:
- vbagr10.chm3077090
ms.prod: excel
api_name:
- Excel.ShowPercentage
ms.assetid: 32e2e547-8fb6-f3c7-3f61-a32a5d77d98d
ms.date: 06/08/2017
---


# ShowPercentage Property

Allows the user to show the percentage value for the data labels on a chart. Read/write Boolean.

 _expression_. **ShowPercentage**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the percentage value to be shown for the data labels of the first series on the first chart.


```vb
Sub UsePercentage() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowPercentage = True 
 
End Sub
```


