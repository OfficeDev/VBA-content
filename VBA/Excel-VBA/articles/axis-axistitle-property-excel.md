---
title: Axis.AxisTitle Property (Excel)
keywords: vbaxl10.chm561075
f1_keywords:
- vbaxl10.chm561075
ms.prod: excel
api_name:
- Excel.Axis.AxisTitle
ms.assetid: 33ba6b94-189b-e9d0-a153-af028380a58a
ms.date: 06/08/2017
---


# Axis.AxisTitle Property (Excel)

Returns an  **[AxisTitle](axistitle-object-excel.md)** object that represents the title of the specified axis. Read-only.


## Syntax

 _expression_ . **AxisTitle**

 _expression_ A variable that represents an **Axis** object.


## Remarks

This example adds an axis label to the category axis in Chart1.


## Example


```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

