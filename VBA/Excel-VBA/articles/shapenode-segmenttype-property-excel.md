---
title: ShapeNode.SegmentType Property (Excel)
keywords: vbaxl10.chm111005
f1_keywords:
- vbaxl10.chm111005
ms.prod: excel
api_name:
- Excel.ShapeNode.SegmentType
ms.assetid: 716e8171-1fd6-941e-209f-e48f5468940f
ms.date: 06/08/2017
---


# ShapeNode.SegmentType Property (Excel)

Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns  **msoSegmentCurve** . Read-only **MsoSegmentType** .


## Syntax

 _expression_ . **SegmentType**

 _expression_ A variable that represents a **ShapeNode** object.


## Remarks



| **MsoSegmentType** can be one of these **MsoSegmentType** constants.|
| **msoSegmentCurve**|
| **msoSegmentLine**|
Use the  **[SetSegmentType](shapenodes-setsegmenttype-method-excel.md)** method to set the value of this property.


## Example

This example changes all straight segments to curved segments in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 n = 1 
 While n <= .Count 
 If .Item(n).SegmentType = msoSegmentLine Then 
 .SetSegmentType n, msoSegmentCurve 
 End If 
 n = n + 1 
 Wend 
End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-excel.md)

