---
title: Shape.Line Property (Excel)
keywords: vbaxl10.chm636101
f1_keywords:
- vbaxl10.chm636101
ms.prod: excel
api_name:
- Excel.Shape.Line
ms.assetid: 0db51c52-c77c-9c0d-9945-e467dbcce3a9
ms.date: 06/08/2017
---


# Shape.Line Property (Excel)

Returns a  **[LineFormat](lineformat-object-excel.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border). Read-only.


## Syntax

 _expression_ . **Line**

 _expression_ A variable that represents a **Shape** object.


## Example

This example adds a blue dashed line to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

This example adds a cross to  `myDocument` and then sets its border to be 8 points thick and red.




```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

