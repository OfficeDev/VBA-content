---
title: ErrorBars Object
keywords: vbagr10.chm131213
f1_keywords:
- vbagr10.chm131213
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: f087bede-5ce2-331f-09e1-4c801f8bca82
ms.date: 06/08/2017
---


# ErrorBars Object

Represents the error bars for the specified chart series. Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups in a 2-D chart can have error bars. Only series in scatter groups can have x and y error bars.

This object isn't a collection. There's no object that represents a single error bar; either you have x error bars or y error bars turned on for all points in a series or you have them turned off.

## Using the ErrorBars Object

Use the  **ErrorBars** property to return the **ErrorBars** object. The following example turns on error bars for series one in `myChart` and then sets the end style for the error bars.


```vb
myChart.SeriesCollection(1).HasErrorBars = True 
myChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```


## Remarks

The  **[ErrorBar](errorbar-method.md)** method changes the format and type of error bars.


