---
title: ShapeNodes.SetPosition Method (Excel)
keywords: vbaxl10.chm112010
f1_keywords:
- vbaxl10.chm112010
ms.prod: excel
api_name:
- Excel.ShapeNodes.SetPosition
ms.assetid: ad76e3d9-51d2-51fd-2af1-9eee7b62e52c
ms.date: 06/08/2017
---


# ShapeNodes.SetPosition Method (Excel)

Sets the location of the node specified by  _Index_. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.


## Syntax

 _expression_ . **SetPosition**( **_Index_** , **_X1_** , **_Y1_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose position is to be set.|
| _X1_|Required| **Single**|The position (in points) of the new node relative to the upper-left corner of the document.|
| _Y1_|Required| **Single**|The position (in points) of the new node relative to the upper-left corner of the document.|

## Example

This example moves node two in shape three on  `myDocument` to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 pointsArray = .Item(2).Points 
 currXvalue = pointsArray(0, 0) 
 currYvalue = pointsArray(0, 1) 
 .SetPosition 2, currXvalue + 200, currYvalue + 300 
End With
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-excel.md)

