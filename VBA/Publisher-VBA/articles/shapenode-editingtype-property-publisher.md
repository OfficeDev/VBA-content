---
title: ShapeNode.EditingType Property (Publisher)
keywords: vbapb10.chm3539200
f1_keywords:
- vbapb10.chm3539200
ms.prod: publisher
api_name:
- Publisher.ShapeNode.EditingType
ms.assetid: f01db634-b35a-48cd-851d-418848674686
ms.date: 06/08/2017
---


# ShapeNode.EditingType Property (Publisher)

If the specified node is a vertex, this property returns an  **MsoEditingType** constant indicating how changes made to the node affect the two segments connected to the node. If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. Read-only.


## Syntax

 _expression_. **EditingType**

 _expression_A variable that represents an  **ShapeNode** object.


### Return Value

MsoEditingType


## Remarks

Use the  **[SetEditingType](shapenodes-seteditingtype-method-publisher.md)** method to set the value of this property.

The  **EditingType** property value can be one of the ** [MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example changes all corner nodes to smooth curve nodes in the third shape in the active publication. The shape must be a freeform drawing.


```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType Index:=intNode, _ 
 EditingType:=msoEditingSmooth 
 End If 
 Next 
End With 

```


