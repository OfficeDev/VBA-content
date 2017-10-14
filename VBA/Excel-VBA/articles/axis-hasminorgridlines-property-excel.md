---
title: Axis.HasMinorGridlines Property (Excel)
keywords: vbaxl10.chm561082
f1_keywords:
- vbaxl10.chm561082
ms.prod: excel
api_name:
- Excel.Axis.HasMinorGridlines
ms.assetid: 27b07e71-448d-33d1-cc4b-472eba7e15d6
ms.date: 06/08/2017
---


# Axis.HasMinorGridlines Property (Excel)

 **True** if the axis has minor gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasMinorGridlines**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the color of the minor gridlines for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 4 
 'set color to green 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

