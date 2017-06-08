---
title: ChartTitle Object
keywords: vbagr10.chm131081
f1_keywords:
- vbagr10.chm131081
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: 6eca7bbc-0158-f25e-d7c8-3f57f06ccccf
ms.date: 06/08/2017
---


# ChartTitle Object

Represents the title of the specified chart.


## Using the ChartTitle Object

Use the  **ChartTitle** property to return the **ChartTitle** object. The following example adds a title to the chart.


```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Remarks

The  **ChartTitle** object doesn't exist and cannot be used unless the **[HasTitle](hastitle-property.md)** property for the chart is  **True**.


