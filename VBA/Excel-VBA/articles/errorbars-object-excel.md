---
title: ErrorBars Object (Excel)
keywords: vbaxl10.chm623072
f1_keywords:
- vbaxl10.chm623072
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: 646de974-bf6f-99c8-20dd-9ca514b7a304
ms.date: 06/08/2017
---


# ErrorBars Object (Excel)

Represents the error bars on a chart series.


## Remarks

 Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2-D chart can have error bars. Only series in scatter groups can have x and y error bars. This object isn't a collection. There's no object that represents a single error bar; you either have x error bars or y error bars turned on for all points in a series or you have them turned off.

The  **[ErrorBar](series-errorbar-method-excel.md)** method changes the error bar format and type.


## Example

Use the  **[ErrorBars](series-errorbars-property-excel.md)** property to return the **ErrorBars** object. The following example turns on error bars for series one in embedded chart one and then sets the end style for the error bars.


```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection(1).HasErrorBars = True 
ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

