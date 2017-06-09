---
title: ShapeNode.Points Property (Excel)
keywords: vbaxl10.chm111004
f1_keywords:
- vbaxl10.chm111004
ms.prod: excel
api_name:
- Excel.ShapeNode.Points
ms.assetid: fe09c78f-44c9-4e66-df7b-c23720216ec5
ms.date: 06/08/2017
---


# ShapeNode.Points Property (Excel)

Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Read-only  **Variant** .


## Syntax

 _expression_ . **Points**

 _expression_ An expression that returns a **ShapeNode** object.


### Return Value

Variant


## Remarks

This property is read-only. Use the  **[SetPosition](shapenodes-setposition-method-excel.md)** method to set the value of this property.


## Example

This example moves node two in shape three on  `myDocument` to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 pointsArray = .Item(2).Points 
 currXvalue = pointsArray(1, 1) 
 currYvalue = pointsArray(1, 2) 
 .SetPosition 2, currXvalue + 200, currYvalue + 300 
End With
```


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)
[ShapeNode Object](shapenode-object-excel.md)

