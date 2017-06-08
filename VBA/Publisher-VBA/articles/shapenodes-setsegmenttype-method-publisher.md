---
title: ShapeNodes.SetSegmentType Method (Publisher)
keywords: vbapb10.chm3473429
f1_keywords:
- vbapb10.chm3473429
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.SetSegmentType
ms.assetid: 64f742fb-8216-9ec3-3fa9-ca2b319cf3e9
ms.date: 06/08/2017
---


# ShapeNodes.SetSegmentType Method (Publisher)

Sets the segment type of the segment that follows the specified node. If the node is a control point for a curved segment, this method sets the segment type for that curve; this may affect the total number of nodes by inserting or deleting adjacent nodes.


## Syntax

 _expression_. **SetSegmentType**( **_Index_**,  **_SegmentType_**)

 _expression_A variable that represents a  **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The node whose segment type is to be set. Must be a number from 1 to the number of nodes in the specified shape; otherwise, an error occurs.|
|SegmentType|Required| **MsoSegmentType**|Specifies the segment type.|

## Remarks

The SegmentType parameter can be one of the  **MsoSegmentType** constants declared in the Microsoft Office type library and shown in the following table.



| **msoSegmentCurve**|
| **msoSegmentLine**|

## Example

This example changes all straight segments to curved segments in the third shape in the active publication. The shape must be a freeform drawing.


```vb
Dim intCount As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 intCount = 1 
 Do While intCount <= .Count 
 If .Item(intCount).SegmentType = msoSegmentLine Then 
 .SetSegmentType _ 
 Index:=intCount, SegmentType:=msoSegmentCurve 
 End If 
 intCount = intCount + 1 
 Loop 
End With 

```


