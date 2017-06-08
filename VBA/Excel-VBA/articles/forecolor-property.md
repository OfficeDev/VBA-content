---
title: ForeColor Property
keywords: vbagr10.chm5207390
f1_keywords:
- vbagr10.chm5207390
ms.prod: excel
api_name:
- Excel.ForeColor
ms.assetid: 1c1eb700-672e-095d-826c-28cdb7e9de40
ms.date: 06/08/2017
---


# ForeColor Property

Returns a  **ChartColorFormat** object that represents the foreground fill color.


## Example

This example sets the gradient, background color, and foreground color for the chart area fill.


```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


