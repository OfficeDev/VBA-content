---
title: Axis.HasMajorGridlines Property (Excel)
keywords: vbaxl10.chm561081
f1_keywords:
- vbaxl10.chm561081
ms.prod: excel
api_name:
- Excel.Axis.HasMajorGridlines
ms.assetid: 2cf9242a-79c5-8288-b71b-a5cd47d5abde
ms.date: 06/08/2017
---


# Axis.HasMajorGridlines Property (Excel)

 **True** if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasMajorGridlines**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the color of the major gridlines for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 3 'set color to red 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

