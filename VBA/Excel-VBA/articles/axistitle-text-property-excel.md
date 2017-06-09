---
title: AxisTitle.Text Property (Excel)
keywords: vbaxl10.chm565085
f1_keywords:
- vbaxl10.chm565085
ms.prod: excel
api_name:
- Excel.AxisTitle.Text
ms.assetid: 1305fae5-afd9-dd8e-f559-f0c6ebff7a3b
ms.date: 06/08/2017
---


# AxisTitle.Text Property (Excel)

Returns or sets the text for the specified object. Read/write  **String** .


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents an **AxisTitle** object.


## Example

This example sets the axis title text for the category axis in Chart1.


```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
End With
```


## See also


#### Concepts


[AxisTitle Object](axistitle-object-excel.md)

