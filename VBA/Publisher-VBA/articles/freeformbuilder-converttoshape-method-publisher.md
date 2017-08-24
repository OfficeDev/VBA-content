---
title: FreeformBuilder.ConvertToShape Method (Publisher)
keywords: vbapb10.chm3276817
f1_keywords:
- vbapb10.chm3276817
ms.prod: publisher
api_name:
- Publisher.FreeformBuilder.ConvertToShape
ms.assetid: 1cb490af-40be-b03f-2f8d-04b1015fbde3
ms.date: 06/08/2017
---


# FreeformBuilder.ConvertToShape Method (Publisher)

Creates a shape that has the geometric characteristics of the specified  **[FreeformBuilder](freeformbuilder-object-publisher.md)** object. Returns a **[Shape](shape-object-publisher.md)** object that represents the new shape.


## Syntax

 _expression_. **ConvertToShape**

 _expression_A variable that represents a  **FreeformBuilder** object.


### Return Value

Shape


## Remarks

You must apply the  **[AddNodes](freeformbuilder-addnodes-method-publisher.md)** method to a  **FreeformBuilder** object at least once before you use the **ConvertToShape** method or an error occurs.


## Example

This example adds a freeform with four vertices to the first page in the active publication.


```vb
' Add a new freeform object. 
With ActiveDocument.Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 

```


