---
title: ShapeNodes.SetPosition Method (Word)
keywords: vbawd10.chm164495374
f1_keywords:
- vbawd10.chm164495374
ms.prod: word
api_name:
- Word.ShapeNodes.SetPosition
ms.assetid: 0675ff22-1717-5fc6-2c07-c7ac53196c88
ms.date: 06/08/2017
---


# ShapeNodes.SetPosition Method (Word)

Sets the location of the node specified by Index.


## Syntax

 _expression_ . **SetPosition**( **_Index_** , **_X1_** , **_Y1_** )

 _expression_ Required. A variable that represents a **[ShapeNodes](shapenodes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose position is to be set.|
| _X1_|Required| **Single**|The position (in points) of the new node relative to the upper-left corner of the document.|

## Remarks

Depending on the editing type of the node, this method may affect the position of adjacent nodes.


## Example

This example moves node two in the third shape on the active document to the right 200 points and down 300 points. The third shape must be a freeform drawing.


```vb
With ActiveDocument.Shapes(3).Nodes 
 pointsArray = .Item(2).Points 
 currXvalue = pointsArray(1, 1) 
 currYvalue = pointsArray(1, 2) 
 .SetPosition 2, currXvalue + 200, currYvalue + 300 
End With
```


## See also


#### Concepts


[ShapeNodes Collection Object](shapenodes-object-word.md)

