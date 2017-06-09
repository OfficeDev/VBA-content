---
title: ShapeNodes.SetEditingType Method (PowerPoint)
keywords: vbapp10.chm560007
f1_keywords:
- vbapp10.chm560007
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.SetEditingType
ms.assetid: ae048107-b416-53f3-ad8b-11a917f7e3dc
ms.date: 06/08/2017
---


# ShapeNodes.SetEditingType Method (PowerPoint)

Sets the editing type of the specified node.


## Syntax

 _expression_. **SetEditingType**( **_Index_**, **_EditingType_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The node whose editing type is to be set.|
| _EditingType_|Required|**MsoEditingType**|The editing type.|

## Remarks

 If the node specified by _Index_ is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.

The  _EditingType_ parameter value can be one of these **MsoEditingType** constants.


||
|:-----|
|**msoEditingAuto**|
|**msoEditingCorner**|
|**msoEditingSmooth**|
|**msoEditingSymmetric**|

## Example

This example changes all corner nodes to smooth nodes in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    For n = 1 to .Count

        If .Item(n).EditingType = msoEditingCorner Then

            .SetEditingType n, msoEditingSmooth

        End If

    Next

End With
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-powerpoint.md)

