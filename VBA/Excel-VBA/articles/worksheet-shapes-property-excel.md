---
title: Worksheet.Shapes Property (Excel)
keywords: vbaxl10.chm174098
f1_keywords:
- vbaxl10.chm174098
ms.prod: excel
api_name:
- Excel.Worksheet.Shapes
ms.assetid: 6206b5e8-742d-797f-12ee-daf3225a53dc
ms.date: 06/08/2017
---


# Worksheet.Shapes Property (Excel)

Returns a  **[Shapes](shapes-object-excel.md)** collection that represents all the shapes on the worksheet. Read-only.


## Syntax

 _expression_ . **Shapes**

 _expression_ An expression that returns a **Worksheet** object.


## Example

This example adds a blue dashed line to worksheet one.


```vb
With Worksheets(1).Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

