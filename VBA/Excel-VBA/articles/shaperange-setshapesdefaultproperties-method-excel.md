---
title: ShapeRange.SetShapesDefaultProperties Method (Excel)
keywords: vbaxl10.chm640093
f1_keywords:
- vbaxl10.chm640093
ms.prod: excel
api_name:
- Excel.ShapeRange.SetShapesDefaultProperties
ms.assetid: 0ddbcaed-827c-5391-db0e-fc1cd94d7b33
ms.date: 06/08/2017
---


# ShapeRange.SetShapesDefaultProperties Method (Excel)

Makes the formatting of the specified shape the default formatting for the shape.


## Syntax

 _expression_ . **SetShapesDefaultProperties**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example adds a rectangle to  `myDocument`, formats the rectangle's fill, sets the rectangle's formatting as the default shape formatting, and then adds another smaller rectangle to the document. The second rectangle has the same fill as the first one.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 With .AddShape(msoShapeRectangle, 5, 5, 80, 60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned msoPatternHorizontalBrick 
 End With 
 ' Set formatting as default formatting 
 .SetShapesDefaultProperties 
 End With 
 ' Create new shape with default formatting 
 .AddShape msoShapeRectangle, 90, 90, 40, 30 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

