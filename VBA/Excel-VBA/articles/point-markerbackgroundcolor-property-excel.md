---
title: Point.MarkerBackgroundColor Property (Excel)
keywords: vbaxl10.chm576084
f1_keywords:
- vbaxl10.chm576084
ms.prod: excel
api_name:
- Excel.Point.MarkerBackgroundColor
ms.assetid: a283c8d2-08f2-0865-b8fe-26bc45d497d8
ms.date: 06/08/2017
---


# Point.MarkerBackgroundColor Property (Excel)

Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .


## Syntax

 _expression_ . **MarkerBackgroundColor**

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

