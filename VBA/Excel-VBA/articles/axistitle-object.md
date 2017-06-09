---
title: AxisTitle Object
keywords: vbagr10.chm131082
f1_keywords:
- vbagr10.chm131082
ms.prod: excel
api_name:
- Excel.AxisTitle
ms.assetid: a5a62dd3-5859-6f5c-5e28-6adbf400e08e
ms.date: 06/08/2017
---


# AxisTitle Object

Represents the title of an axis in a chart.


## Using the AxisTitle Object

Use the  **AxisTitle** property to return an **AxisTitle** object. The following example sets the text of the value axis title and sets the font to 10-point Bookman.


```vb
With myChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 End With 
End With
```


## Remarks

The  **AxisTitle** object doesn't exist and cannot be used unless the **[HasTitle](hastitle-property.md)** property for the specified axis is  **True**.


