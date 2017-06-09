---
title: FreeformBuilder.AddNodes Method (Word)
keywords: vbawd10.chm164167690
f1_keywords:
- vbawd10.chm164167690
ms.prod: word
api_name:
- Word.FreeformBuilder.AddNodes
ms.assetid: 793e869f-2365-1ef0-f2e4-d764f67f0cb9
ms.date: 06/08/2017
---


# FreeformBuilder.AddNodes Method (Word)

Inserts a new segment at the end of the freeform that's being created, and adds the nodes that define the segment.


## Syntax

 _expression_ . **AddNodes**( **_SegmentType_** , **_EditingType_** , **_X1_** , **_Y1_** , **_X2_** , **_Y2_** , **_X3_** , **_Y3_** )

 _expression_ Required. A variable that represents a **[FreeformBuilder](freeformbuilder-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SegmentType_|Required| **MsoSegmentType**|The type of segment to be added.|
| _EditingType_|Required| **MsoEditingType**|The editing property of the vertex. If SegmentType is  **msoSegmentLine** , EditingType must be **msoEditingAuto** .|
| _X1_|Required| **Single**|If the EditingType of the new segment is  **msoEditingAuto** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the EditingType of the new segment is  **msoEditingAuto** , this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner** , this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y2_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto** , do not specify a value for this argument.|
| _X3_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y3_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|

## Remarks

You can use this method as many times as you want to add nodes to the freeform you are creating. When you finish adding nodes, use the  **ConvertToShape** method to create the freeform you've just defined. To add nodes to a freeform after it has been created, use the **Insert** method of the **[ShapeNodes](shapenodes-object-word.md)** collection.


## Example

This example adds a freeform with five vertices to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
 .AddNodes msoSegmentCurve, msoEditingCorner, _ 
 380, 230, 400, 250, 450, 300 
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
 .ConvertToShape 
End With
```


## See also


#### Concepts


[FreeformBuilder Object](freeformbuilder-object-word.md)

