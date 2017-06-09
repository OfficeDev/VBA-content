---
title: ShapeNodes.SetEditingType Method (Publisher)
keywords: vbapb10.chm3473427
f1_keywords:
- vbapb10.chm3473427
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.SetEditingType
ms.assetid: f90b1323-d682-1b2b-6747-cea5f2cead3c
ms.date: 06/08/2017
---


# ShapeNodes.SetEditingType Method (Publisher)

Sets the editing type of the specified node. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Depending on the editing type, this method may affect the position of adjacent nodes.


## Syntax

 _expression_. **SetEditingType**( **_Index_**,  **_EditingType_**)

 _expression_A variable that represents a  **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The node whose editing type is to be set. Must be a number from 1 to the number of nodes in the specified shape; otherwise, an error occurs.|
|EditingType|Required| **MsoEditingType**|The editing property of the node.|

## Remarks

The EditingType parameter can be one of the  **MsoEditingType** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoEditingAuto**|Changes the node to a type appropriate to the segments being connected.|
| **msoEditingCorner**| Changes the node to a corner node.|
| **msoEditingSmooth**|Changes the node to a smooth curve node..|
| **msoEditingSymmetric**|Changes the node to a symmetric curve node.|

## Example

This example changes all corner nodes to smooth nodes in the third shape of the active publication. The shape must be a freeform drawing.


```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType _ 
 Index:=intNode, EditingType:=msoEditingSmooth 
 End If 
 Next intNode 
End With 

```


