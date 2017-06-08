---
title: ShapeRange.ThreeD Property (Excel)
keywords: vbaxl10.chm640116
f1_keywords:
- vbaxl10.chm640116
ms.prod: excel
api_name:
- Excel.ShapeRange.ThreeD
ms.assetid: 0b4ab4b8-841b-eea6-67a4-effe144d19fe
ms.date: 06/08/2017
---


# ShapeRange.ThreeD Property (Excel)

Returns a  **[ThreeDFormat](threedformat-object-excel.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **ThreeD**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3-D effects applied to shape one on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 ' RGB value for purple 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

