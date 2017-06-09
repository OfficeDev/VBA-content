---
title: ShapeNodes.SetEditingType Method (Word)
keywords: vbawd10.chm164495373
f1_keywords:
- vbawd10.chm164495373
ms.prod: word
api_name:
- Word.ShapeNodes.SetEditingType
ms.assetid: 315a8a0d-0caa-278d-af0e-91b468b694ab
ms.date: 06/08/2017
---


# ShapeNodes.SetEditingType Method (Word)

Sets the editing type of the node specified by Index. .


## Syntax

 _expression_ . **SetEditingType**( **_Index_** , **_EditingType_** )

 _expression_ Required. A variable that represents a **[ShapeNodes](shapenodes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose editing type is to be set.|
| _EditingType_|Required| **MsoEditingType**|The editing property of the vertex.|

## Remarks

If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Depending on the editing type, this method may affect the position of adjacent nodes.


## Example

This example changes all corner nodes to smooth nodes in the third shape on the active document. The third shape must be a freeform drawing.


```vb
Dim lngLoop As lngLoop 
 
With ActiveDocument.Shapes(3).Nodes 
 For lngLoop = 1 to .Count 
 If .Item(lngLoop).EditingType = msoEditingCorner Then 
 .SetEditingType lngLoop, msoEditingSmooth 
 End If 
 Next lngLoop 
End With
```


## See also


#### Concepts


[ShapeNodes Collection Object](shapenodes-object-word.md)

