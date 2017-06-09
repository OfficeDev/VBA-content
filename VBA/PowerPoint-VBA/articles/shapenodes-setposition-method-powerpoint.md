---
title: ShapeNodes.SetPosition Method (PowerPoint)
keywords: vbapp10.chm560008
f1_keywords:
- vbapp10.chm560008
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.SetPosition
ms.assetid: 8defcf80-84d8-538a-2dce-d3ffe5e8dfb0
ms.date: 06/08/2017
---


# ShapeNodes.SetPosition Method (PowerPoint)

Sets the location of the node specified by  **Index**. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.


## Syntax

 _expression_. **SetPosition**( **_Index_**, **_X1_**, **_Y1_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The node whose position is to be set.|
| _ Y1_|Required|**Single**|The x-position (in points) of the new node relative to the upper-left corner of the document.|
| _ Y1_|Required|**Single**|The y-position (in points) of the new node relative to the upper-left corner of the document.|

## Example

This example moves node two in shape three on  `myDocument` to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    pointsArray = .Item(2).Points

    currXvalue = pointsArray(1, 1)

    currYvalue = pointsArray(1, 2)

    .SetPosition 2, currXvalue + 200, currYvalue + 300

End With
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-powerpoint.md)

