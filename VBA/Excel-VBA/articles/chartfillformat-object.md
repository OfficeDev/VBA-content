---
title: ChartFillFormat Object
keywords: vbagr10.chm5207187
f1_keywords:
- vbagr10.chm5207187
ms.prod: excel
api_name:
- Excel.ChartFillFormat
ms.assetid: e011f58f-141b-1b21-0db4-04a5c5e964c6
ms.date: 06/08/2017
---


# ChartFillFormat Object

Represents fill formatting.


## Using the ChartFillFormat Object

Use the  **[Fill](fill-property.md)** property to return the **ChartFillFormat** object. The following example sets the foreground color, background color, and gradient for the chart area fill in `myChart`.


```vb
With myChart.ChartArea.Fill 
    .Visible = True 
    .ForeColor.SchemeColor = 15 
    .BackColor.SchemeColor = 17 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```


