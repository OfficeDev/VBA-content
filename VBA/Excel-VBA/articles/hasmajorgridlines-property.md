---
title: HasMajorGridlines Property
keywords: vbagr10.chm5207498
f1_keywords:
- vbagr10.chm5207498
ms.prod: excel
api_name:
- Excel.HasMajorGridlines
ms.assetid: f3c22d5d-4150-43b1-5f0d-3d49049e1e24
ms.date: 06/08/2017
---


# HasMajorGridlines Property

 **True** if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean**.


## Example

This example sets the color of the major gridlines for the value axis.


```vb
With myChart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 3 'set color to red 
 End If 
End With
```


