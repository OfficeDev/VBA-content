---
title: ShapeNodes.SetSegmentType Method (PowerPoint)
keywords: vbapp10.chm560009
f1_keywords:
- vbapp10.chm560009
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.SetSegmentType
ms.assetid: 8dfca78c-db97-b0a5-37e9-232354c2e21f
ms.date: 06/08/2017
---


# ShapeNodes.SetSegmentType Method (PowerPoint)

Sets the segment type of the segment that follows the specified node.


## Syntax

 _expression_. **SetSegmentType**( **_Index_**, **_SegmentType_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The node whose segment type is to be set.|
| _SegmentType_|Required|**MsoSegmentType**|Specifies if the segment is straight or curved.|

## Remarks

 If the node specified by Index is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.

The  _SegmentType_ parameter value can be one of these **MsoSegmentType** constants.


||
|:-----|
|**msoSegmentCurve**|
|**msoSegmentLine**|

## Example

This example changes all straight segments to curved segments in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

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


[ShapeNodes Object](shapenodes-object-powerpoint.md)

