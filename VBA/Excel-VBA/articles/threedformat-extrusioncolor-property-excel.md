---
title: ThreeDFormat.ExtrusionColor Property (Excel)
keywords: vbaxl10.chm119006
f1_keywords:
- vbaxl10.chm119006
ms.prod: excel
api_name:
- Excel.ThreeDFormat.ExtrusionColor
ms.assetid: d9c76fe5-69dc-5bdd-8882-7f06ba083947
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColor Property (Excel)

Returns a  **[ColorFormat](colorformat-object-excel.md)** object that represents the color of the shape's extrusion. Read-only.


## Syntax

 _expression_ . **ExtrusionColor**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Example

This example adds an oval to  `myDocument` and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Set myDocument = Worksheets(1) 
Set myShape = myDocument.Shapes.AddShape(msoShapeOval, _ 
 90, 90, 90, 40) 
With myShape.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 ' RGB value for purple 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

