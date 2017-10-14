---
title: ShapeNodes.Insert Method (Word)
keywords: vbawd10.chm164495372
f1_keywords:
- vbawd10.chm164495372
ms.prod: word
api_name:
- Word.ShapeNodes.Insert
ms.assetid: a0a8a577-b0c5-fad8-da21-f3adbbdde085
ms.date: 06/08/2017
---


# ShapeNodes.Insert Method (Word)

Inserts a node into a freeform shape.


## Syntax

 _expression_ . **Insert**( **_Index_** , **_SegmentType_** , **_EditingType_** , **_X1_** , **_Y1_** , **_X2_** , **_Y2_** , **_X3_** , **_Y3_** )

 _expression_ Required. A variable that represents a **[ShapeNodes](shapenodes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The number of the shape node after which to insert a new node.|
| _SegmentType_|Required| **MsoSegmentType**|The type of line that connects the inserted node to the neighboring nodes.|
| _EditingType_|Required| **MsoEditingType**|The editing property of the inserted node.|
| _X1_|Required| **Single**|If the EditingType of the new segment is  **msoEditingAuto** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the starting point of the new segment. If the EditingType of the new node is **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the EditingType of the new segment is  **msoEditingAuto** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the starting point of the new segment. If the EditingType of the new node is **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y2_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _X3_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the ending point of the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y3_|Optional| **Single**|If the EditingType of the new segment is  **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the ending point of the new segment. If the EditingType of the new segment is **msoEditingAuto** , don't specify a value for this argument.|

## Example

This example selects the third shape in the active document, checks whether the shape is a  **Freeform** object, and if it is, inserts a node.


```vb
Sub InsertShapeNode() 
 ActiveDocument.Shapes(3).Select 
 With Selection.ShapeRange 
 If .Type = msoFreeform Then 
 .Nodes.Insert _ 
 Index:=3, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSymmetric, x1:=35, y1:=100 
 .Fill.ForeColor.RGB = RGB(0, 0, 200) 
 .Fill.Visible = msoTrue 
 Else 
 MsgBox "This shape is not a Freeform object." 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[ShapeNodes Collection Object](shapenodes-object-word.md)

