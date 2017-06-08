---
title: MajorGridlines Property
keywords: vbagr10.chm65625
f1_keywords:
- vbagr10.chm65625
ms.prod: excel
api_name:
- Excel.MajorGridlines
ms.assetid: d160f530-e92e-4528-e207-d47ae710a7d5
ms.date: 06/08/2017
---


# MajorGridlines Property

Returns a  **[Gridlines](gridlines-object.md)** object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Example

This example sets the color of the major gridlines for the value axis in the chart.


```vb
With myChart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 5 
 End If 
End With
```


