---
title: Point.MarkerForegroundColorIndex Property (Excel)
keywords: vbaxl10.chm576087
f1_keywords:
- vbaxl10.chm576087
ms.prod: excel
api_name:
- Excel.Point.MarkerForegroundColorIndex
ms.assetid: 00d5e240-0851-ea13-11a3-5972135ca5fa
ms.date: 06/08/2017
---


# Point.MarkerForegroundColorIndex Property (Excel)

Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Applies only to line, scatter, and radar charts. Read/write **Long** .


## Syntax

 _expression_ . **MarkerForegroundColorIndex**

 _expression_ A variable that represents a **Point** object.


## Example

This example sets the marker background and foreground colors for the second point in series one in Chart1.


```vb
With Charts("Chart1").SeriesCollection(1).Points(2) 
 .MarkerBackgroundColorIndex = 4 'green 
 .MarkerForegroundColorIndex = 3 'red 
End With
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

