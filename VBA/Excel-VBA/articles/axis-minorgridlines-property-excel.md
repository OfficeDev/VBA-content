---
title: Axis.MinorGridlines Property (Excel)
keywords: vbaxl10.chm561092
f1_keywords:
- vbaxl10.chm561092
ms.prod: excel
api_name:
- Excel.Axis.MinorGridlines
ms.assetid: 5725fdb3-05de-e555-5734-cbc64c6a2068
ms.date: 06/08/2017
---


# Axis.MinorGridlines Property (Excel)

Returns a  **[Gridlines](gridlines-object-excel.md)** object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Syntax

 _expression_ . **MinorGridlines**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the color of the minor gridlines for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 5 'set color to blue 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

