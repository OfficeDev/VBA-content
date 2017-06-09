---
title: MarkerForegroundColor Property
keywords: vbagr10.chm5207661
f1_keywords:
- vbagr10.chm5207661
ms.prod: excel
api_name:
- Excel.MarkerForegroundColor
ms.assetid: 27c88341-0446-bad5-25f4-a4f19c2db4ec
ms.date: 06/08/2017
---


# MarkerForegroundColor Property

Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts. Read/write  **Long**.


## Example

This example sets the marker background and foreground colors for the second point in series one.


```vb
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```


