---
title: DocumentWindow.PointsToScreenPixelsX Method (PowerPoint)
keywords: vbapp10.chm511027
f1_keywords:
- vbapp10.chm511027
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.PointsToScreenPixelsX
ms.assetid: 6b5f2f58-41af-3620-74f3-1c4ec3922fc2
ms.date: 06/08/2017
---


# DocumentWindow.PointsToScreenPixelsX Method (PowerPoint)

Converts a horizontal measurement from points to pixels. Used to return a horizontal screen location for a text frame or shape. Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToScreenPixelsX**( **_Points_** )

 _expression_ A variable that represents a **DocumentWindow** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required|**Single**|The horizontal measurement (in points) to be converted to pixels.|

### Return Value

Single


## Example

This example converts the width and height of the selected text frame bounding box from points to pixels, and returns the values to  `myXparm` and `myYparm`.


```vb
With ActiveWindow
    myXparm = .PointsToScreenPixelsX _
        (.Selection.TextRange.BoundWidth)
    myYparm = .PointsToScreenPixelsY _
        (.Selection.TextRange.BoundHeight)
End With
```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

