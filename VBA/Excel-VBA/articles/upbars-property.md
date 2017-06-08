---
title: UpBars Property
keywords: vbagr10.chm5208103
f1_keywords:
- vbagr10.chm5208103
ms.prod: excel
api_name:
- Excel.UpBars
ms.assetid: e0a27db4-276c-446d-af89-b3b9aa962412
ms.date: 06/08/2017
---


# UpBars Property

Returns an  **[UpBars](upbars-object.md)** object that represents the up bars on a line chart. Applies only to line charts. Read-only.


## Example

This example turns on up and down bars for chart group one and then sets their colors. The example should be run on a 2-D line chart containing two series that cross each other at one or more data points.


```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```


