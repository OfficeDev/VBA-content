---
title: ThreeDFormat.Depth Property (Excel)
keywords: vbaxl10.chm119005
f1_keywords:
- vbaxl10.chm119005
ms.prod: excel
api_name:
- Excel.ThreeDFormat.Depth
ms.assetid: 1fce69d1-6813-1f92-d457-6a6c36de7dba
ms.date: 06/08/2017
---


# ThreeDFormat.Depth Property (Excel)

Returns or sets a  **Single** value that represents the depth of the shape's extrusion.


## Syntax

 _expression_ . **Depth**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

The value of this property can be a value from -600 through 9600 (positive values produce an extrusion whose front face is the original shape; negative values produce an extrusion whose back face is the original shape).


## Example

This example adds an oval to myDocument and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Set myDocument = Worksheets(1) 
Set myShape = myDocument.Shapes.AddShape(msoShapeOval, _ 
 90, 90, 90, 40) 
With myShape.ThreeD 
 .Visible = True 
 .Depth = 50 
 ' RGB value for purple 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

