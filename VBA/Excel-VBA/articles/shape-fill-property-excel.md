---
title: Shape.Fill Property (Excel)
keywords: vbaxl10.chm636096
f1_keywords:
- vbaxl10.chm636096
ms.prod: excel
api_name:
- Excel.Shape.Fill
ms.assetid: b533b463-51c5-f59e-c3ba-cfe7512daa53
ms.date: 06/08/2017
---


# Shape.Fill Property (Excel)

Returns a  **[FillFormat](fillformat-object-excel.md)** object for a specified shape or a **[ChartFillFormat](http://msdn.microsoft.com/library/e011f58f-141b-1b21-0db4-04a5c5e964c6%28Office.15%29.aspx)** object for a specified chart that contains fill formatting properties for the shape or chart. Read-only.


## Syntax

 _expression_ . **Fill**

 _expression_ A variable that represents a **Shape** object.


## Example

This example adds a rectangle to  `myDocument` and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
        90, 90, 90, 50).Fill 
    .ForeColor.RGB = RGB(128, 0, 0) 
    .BackColor.RGB = RGB(170, 170, 170) 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

