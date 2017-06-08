---
title: HasTitle Property
keywords: vbagr10.chm65590
f1_keywords:
- vbagr10.chm65590
ms.prod: excel
api_name:
- Excel.HasTitle
ms.assetid: 9ecc48d3-fd86-e185-a69f-0676230b3194
ms.date: 06/08/2017
---


# HasTitle Property

 **True** if the axis or chart has a visible title. Read/write **Boolean**.


## Remarks

An axis title is represented by an  **[AxisTitle](axistitle-object.md)** object.

A chart title is represented by a  **[ChartTitle](charttitle-object.md)** object.


## Example

This example adds an axis label to the category axis.


```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```


