---
title: Point.MarkerForegroundColor Property (Excel)
keywords: vbaxl10.chm576086
f1_keywords:
- vbaxl10.chm576086
ms.prod: excel
api_name:
- Excel.Point.MarkerForegroundColor
ms.assetid: 800fb100-8dc3-8e03-7308-48ffb2df552e
ms.date: 06/08/2017
---


# Point.MarkerForegroundColor Property (Excel)

Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .


## Syntax

 _expression_ . **MarkerForegroundColor**

 _expression_ A variable that represents a **Point** object.


## Example

This example sets the marker background and foreground colors for the second point in series one in Chart1.


```vb
With Charts("Chart1").SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

