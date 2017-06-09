---
title: ShapeNodes.SetSegmentType Method (Word)
keywords: vbawd10.chm164495375
f1_keywords:
- vbawd10.chm164495375
ms.prod: word
api_name:
- Word.ShapeNodes.SetSegmentType
ms.assetid: 8afa8b4b-73bf-e64b-b6fa-427e891a9e07
ms.date: 06/08/2017
---


# ShapeNodes.SetSegmentType Method (Word)

Sets the segment type of the segment that follows the node specified by Index.


## Syntax

 _expression_ . **SetSegmentType**( **_Index_** , **_SegmentType_** )

 _expression_ Required. A variable that represents a **[ShapeNodes](shapenodes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose segment type is to be set.|
| _SegmentType_|Required| **MsoSegmentType**|Specifies if the segment is straight or curved.|

## Remarks

If the node is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.


## Example

This example changes all straight segments to curved segments in the third shape on the active document. The third shape must be a freeform drawing.


```vb
Dim lngLoop As Long 
 
With ActiveDocument.Shapes(3).Nodes 
 lngLoop = 1 
 While lngLoop <= .Count 
 If .Item(lngLoop).SegmentType = msoSegmentLine Then 
 .SetSegmentType lngLoop, msoSegmentCurve 
 End If 
 lngLoop = lngLoop + 1 
 Wend 
End With
```


## See also


#### Concepts


[ShapeNodes Collection Object](shapenodes-object-word.md)

