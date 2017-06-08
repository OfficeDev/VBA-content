---
title: ShapeRange.Nodes Property (Excel)
keywords: vbaxl10.chm640111
f1_keywords:
- vbaxl10.chm640111
ms.prod: excel
api_name:
- Excel.ShapeRange.Nodes
ms.assetid: 6005d3f3-2c08-f539-87fc-51425ce81e0e
ms.date: 06/08/2017
---


# ShapeRange.Nodes Property (Excel)

Returns a  **[ShapeNodes](shapenodes-object-excel.md)** collection that represents the geometric description of the specified shape.


## Syntax

 _expression_ . **Nodes**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

This property applies to  **[Shape](shape-object-excel.md)** or **[ShapeRange](shaperange-object-excel.md)** objects that represent freeform drawings.


## Example

This example adds a smooth node with a curved segment after node four in shape three on  `myDocument`. Shape three must be a freeform drawing with at least four nodes.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

