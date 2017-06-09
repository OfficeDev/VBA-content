---
title: MinorGridlines Property
keywords: vbagr10.chm5207695
f1_keywords:
- vbagr10.chm5207695
ms.prod: excel
api_name:
- Excel.MinorGridlines
ms.assetid: 80ca57a1-7e8f-4d83-0da6-2a28399c27af
ms.date: 06/08/2017
---


# MinorGridlines Property

Returns a  **[Gridlines](gridlines-object.md)** object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Example

This example sets the color of the minor gridlines for the value axis in the chart to blue.


```vb
With myChart.Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 5 
 End If 
End With
```


