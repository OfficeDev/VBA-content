---
title: ErrorBars Property
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: 28e7e234-3731-42b6-b8dc-f1945b30678e
ms.date: 06/08/2017
---


# ErrorBars Property

Returns an  [ErrorBars](errorbars-object.md) object that represents the error bars for the series. Read-only.


## Example

This example sets the error bar color for series one. The example should be run on a 2-D line chart that has error bars for series one.


```vb
With myChart.SeriesCollection(1)
    .ErrorBars.Border.ColorIndex = 8
End With
```


