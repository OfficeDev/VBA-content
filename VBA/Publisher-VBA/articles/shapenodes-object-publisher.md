---
title: ShapeNodes Object (Publisher)
keywords: vbapb10.chm3538943
f1_keywords:
- vbapb10.chm3538943
ms.prod: publisher
api_name:
- Publisher.ShapeNodes
ms.assetid: f190a8a8-e03a-e8a2-482a-5e092ff3ed86
ms.date: 06/08/2017
---


# ShapeNodes Object (Publisher)

A collection of all the  **[ShapeNode](shapenode-object-publisher.md)** objects in the specified freeform. Each **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the **[BuildFreeform](shapes-buildfreeform-method-publisher.md)** and **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** methods.
 


## Example

Use the  **[Nodes](shape-nodes-property-publisher.md)** property to return a **ShapeNodes** collection. The following example deletes node four in shape three on the active document. For this example to work, shape three must be a freeform with at least four nodes.
 

 

```
Sub DeleteShapeNode() 
 ActiveDocument.Pages(1).Shapes(3).Nodes.Delete Index:=4 
End Sub
```

Use the  **[Insert](shapenodes-insert-method-publisher.md)** method to create a new node and add it to the **ShapeNodes** collection. The following example adds a smooth node with a curved segment after node four in shape three on the active document. For this example to work, shape three must e a freeform with at least four nodes.
 

 



```
Sub AddCurvedSmoothSegment() 
 ActiveDocument.Pages(1).Shapes(3).Nodes.Insert _ 
 Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End Sub
```

Use  **Nodes** (index), where index is the node index number, to return a single **ShapeNode** object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.
 

 



```
Sub SetPointType() 
 With ActiveDocument.Pages(1).Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType Index:=1, EditingType:=msoEditingSmooth 
 End If 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](shapenodes-delete-method-publisher.md)|
|[Insert](shapenodes-insert-method-publisher.md)|
|[Item](shapenodes-item-method-publisher.md)|
|[SetEditingType](shapenodes-seteditingtype-method-publisher.md)|
|[SetPosition](shapenodes-setposition-method-publisher.md)|
|[SetSegmentType](shapenodes-setsegmenttype-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](shapenodes-application-property-publisher.md)|
|[Count](shapenodes-count-property-publisher.md)|
|[Parent](shapenodes-parent-property-publisher.md)|

