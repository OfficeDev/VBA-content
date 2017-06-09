---
title: ChartColorFormat Object
keywords: vbagr10.chm131251
f1_keywords:
- vbagr10.chm131251
ms.prod: excel
api_name:
- Excel.ChartColorFormat
ms.assetid: 5d2e0cb0-e928-0704-7b4c-1afee6096f3a
ms.date: 06/08/2017
---


# ChartColorFormat Object

Represents a foreground or background color.


## Using the ChartColorFormat Object

Use the  **[ForeColor](forecolor-property.md)** property to return a **ChartColorFormat** object that represents the foreground fill color. Use the **[BackColor](backcolor-property-graph.md)** property to return the background fill color. Use the **[RGB](rgb-property.md)** property to return the color as an explicit red-green-blue value, and use the **[SchemeColor](schemecolor-property.md)** property to return or set the color as one of the colors in the current color scheme. The following example sets the foreground color, background color, and gradient for the chart area fill in `myChart`.


```vb
With myChart.ChartArea.Fill 
    .Visible = True 
    .ForeColor.SchemeColor = 15 
    .BackColor.SchemeColor = 17 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```


