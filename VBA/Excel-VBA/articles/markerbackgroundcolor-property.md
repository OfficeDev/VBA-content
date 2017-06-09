---
title: MarkerBackgroundColor Property
keywords: vbagr10.chm65609
f1_keywords:
- vbagr10.chm65609
ms.prod: excel
api_name:
- Excel.MarkerBackgroundColor
ms.assetid: 035d3bf9-e6cf-7f43-aaee-fc3c3926afaa
ms.date: 06/08/2017
---


# MarkerBackgroundColor Property

Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. Read/write  **Long**.


## Example

This example sets the marker background and foreground colors for the second point in series one.


```vb
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```


