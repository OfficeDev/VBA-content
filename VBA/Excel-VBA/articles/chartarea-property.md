---
title: ChartArea Property
keywords: vbagr10.chm5207181
f1_keywords:
- vbagr10.chm5207181
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 1af59d11-2b63-d629-5dae-d9b9d8303ddf
ms.date: 06/08/2017
---


# ChartArea Property

Returns a  **[ChartArea](chartarea-object.md)** object that represents the complete chart area for the chart. Read-only.


## Example

This example sets the chart area interior color of  `myChart` to red and sets the border color to blue.


```vb
With myChart.ChartArea 
    .Interior.ColorIndex = 3 
    .Border.ColorIndex = 5 
End With
```


