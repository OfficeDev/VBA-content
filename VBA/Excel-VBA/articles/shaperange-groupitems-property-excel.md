---
title: ShapeRange.GroupItems Property (Excel)
keywords: vbaxl10.chm640104
f1_keywords:
- vbaxl10.chm640104
ms.prod: excel
api_name:
- Excel.ShapeRange.GroupItems
ms.assetid: daf6d12c-409a-cf0a-989f-319333d24596
ms.date: 06/08/2017
---


# ShapeRange.GroupItems Property (Excel)

Returns a  **[GroupShapes](groupshapes-object-excel.md)** object that represents the individual shapes in the specified group. Use the **[Item](groupshapes-item-method-excel.md)** method of the **GroupShapes** object to return a single shape from the group. Applies to **ShapeRange** objects that represent grouped shapes. Read-only.


## Syntax

 _expression_ . **GroupItems**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example adds three triangles to  `myDocument`, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 10, 10, 100, 100).Name = "shpOne" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 150, 10, 100, 100).Name = "shpTwo" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 300, 10, 100, 100).Name = "shpThree" 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble 
 End With 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

