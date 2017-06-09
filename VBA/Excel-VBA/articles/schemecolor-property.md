---
title: SchemeColor Property
keywords: vbagr10.chm5207954
f1_keywords:
- vbagr10.chm5207954
ms.prod: excel
api_name:
- Excel.SchemeColor
ms.assetid: a90b4570-dae3-4ca1-563a-0467efbf9bca
ms.date: 06/08/2017
---


# SchemeColor Property

Returns or sets the color of the specified  **ChartColorFormat** object as an index in the current color scheme. Read/write **Long**.


## Example

This example sets the foreground color, background color, and gradient for the chart area fill on the chart.


```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


