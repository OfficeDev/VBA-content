---
title: FreeformBuilder.AddNodes Method (Publisher)
keywords: vbapb10.chm3276816
f1_keywords:
- vbapb10.chm3276816
ms.prod: publisher
api_name:
- Publisher.FreeformBuilder.AddNodes
ms.assetid: 29906bde-e6a6-f661-0f3f-085f39653e42
ms.date: 06/08/2017
---


# FreeformBuilder.AddNodes Method (Publisher)

Inserts a new segment at the end of the freeform that is being created, and adds the nodes that define the segment. You can use this method as many times as you want to add nodes to the freeform you are creating. When you finish adding nodes, use the  **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** method to create the freeform you just defined.


## Syntax

 _expression_. **AddNodes**( **_SegmentType_**,  **_EditingType_**,  **_X1_**,  **_Y1_**,  **_X2_**,  **_Y2_**,  **_X3_**,  **_Y3_**)

 _expression_A variable that represents a  **FreeformBuilder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SegmentType|Required| **MsoSegmentType**|The type of segment to be added.|
|EditingType|Required| **MsoEditingType**|Specifies the editing type of the new node. If SegmentType is  **msoSegmentLine**, EditingType must be  **msoEditingAuto**; otherwise, an error occurs.|
|X1|Required| **Variant**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new node is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the first control point for the new segment.|
|Y1|Required| **Variant**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new node is  **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the first control point for the new segment.|
|X2|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the second control point for the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|Y2|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the second control point for the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|X3|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|Y3|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|

## Remarks

SegmentType can be one of these  **MsoSegmentType** constants.



| **msoSegmentCurve**|
| **msoSegmentLine**|
EditingType can be one of these  **MsoEditingType** constants.



| **msoEditingAuto**|Adds a node type appropriate to the segments being connected.|
| **msoEditingCorner**|Adds a corner node.|
For the X1, Y1, X2, Y2, X3, and Y3 arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

To add nodes to a freeform after iit is created, use the  **[Insert](shapenodes-insert-method-publisher.md)** method of the  **[ShapeNodes](shapenodes-object-publisher.md)** collection.


## Example

This example adds a freeform with four vertices to the first page in the active publication.


```vb
' Add a new freeform object. 
With ActiveDocument.Pages(1).Shapes _ 
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


