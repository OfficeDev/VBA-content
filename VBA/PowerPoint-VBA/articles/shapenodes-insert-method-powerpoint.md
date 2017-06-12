---
title: ShapeNodes.Insert Method (PowerPoint)
keywords: vbapp10.chm560006
f1_keywords:
- vbapp10.chm560006
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.Insert
ms.assetid: ece6e886-db56-6800-fe1c-f9d308104d75
ms.date: 06/08/2017
---


# ShapeNodes.Insert Method (PowerPoint)

Inserts a new segment after the specified node of the freeform.


## Syntax

 _expression_. **Insert**( **_Index_**, **_SegmentType_**, **_EditingType_**, **_X1_**, **_Y1_**, **_X2_**, **_Y2_**, **_X3_**, **_Y3_** )

 _expression_ A variable that represents an **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The node that the new node is to be inserted after.|
| _SegmentType_|Required|**MsoSegmentType**|The type of segment to be added.|
| _EditingType_|Required|**MsoEditingType**|The editing property of the vertex.|
| _X1_|Required|**Single**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required|**Single**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y2_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _X3_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y3_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|

## Remarks

The  _SegmentType_ parameter value can be one of these **MsoSegmentType** constants.


||
|:-----|
|**msoSegmentCurve**|
|**msoSegmentLine**|
The  _EditingType_ parameter value can be one of these **MsoEditingType** constants.


||
|:-----|
|**msoEditingAuto**|
|**msoEditingCorner**|

## Example

This example adds a smooth node with a curved segment after node four in shape three on  `myDocument`. Shape three must be a freeform drawing with at least four nodes.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes
    .Insert Index:=4, SegmentType:=msoSegmentCurve, _
        EditingType:=msoEditingSmooth, X1:=210, Y1:=100
End With
```


