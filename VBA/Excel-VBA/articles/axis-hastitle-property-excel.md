---
title: Axis.HasTitle Property (Excel)
keywords: vbaxl10.chm561083
f1_keywords:
- vbaxl10.chm561083
ms.prod: excel
api_name:
- Excel.Axis.HasTitle
ms.assetid: 4b3d656f-4416-42a6-cefd-9684ba98c8e3
ms.date: 06/08/2017
---


# Axis.HasTitle Property (Excel)

 **True** if the axis or chart has a visible title. Read/write **Boolean** .


## Syntax

 _expression_ . **HasTitle**

 _expression_ A variable that represents an **Axis** object.


## Remarks

An axis title is represented by an  **[AxisTitle](axistitle-object-excel.md)** object.


## Example

This example adds an axis label to the category axis in Chart1.


```vb
With Charts("Chart1").Axis(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

