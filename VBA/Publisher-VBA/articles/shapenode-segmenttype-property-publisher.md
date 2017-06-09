---
title: ShapeNode.SegmentType Property (Publisher)
keywords: vbapb10.chm3539202
f1_keywords:
- vbapb10.chm3539202
ms.prod: publisher
api_name:
- Publisher.ShapeNode.SegmentType
ms.assetid: 471206b2-ca37-5e4a-678b-df8a47c90f96
ms.date: 06/08/2017
---


# ShapeNode.SegmentType Property (Publisher)

Returns an  **MsoSegmentType** constant that indicates whether the segment associated with the specified node is straight or curved. Read-only.


## Syntax

 _expression_. **SegmentType**

 _expression_A variable that represents a  **ShapeNode** object.


### Return Value

MsoSegmentType


## Remarks

If the specified node is a control point for a curved segment, this property returns  **msoSegmentCurve**.

Use the  **[SetSegmentType](shapenodes-setsegmenttype-method-publisher.md)** method to set the value of this property.

The  **SegmentType** property value can be one of these **MsoSegmentType** constants declared in the Microsoft Publisher type library.



| **msoSegmentCurve**|
| **msoSegmentLine**|

## Example

This example changes all straight segments to curved segments in the first shape on the first page of the active publication. For this example to work, the specified shape must be a freeform drawing.


```vb
Sub ChangeSegmentTypes() 
 Dim intNode As Integer 
 With ActiveDocument.Pages(1).Shapes(1).Nodes 
 intNode = 1 
 Do While intNode <= .Count 
 If .Item(intNode).SegmentType = msoSegmentLine Then 
 .SetSegmentType Index:=intNode, _ 
 SegmentType:=msoSegmentCurve 
 End If 
 intNode = intNode + 1 
 Loop 
 End With 
End Sub
```


