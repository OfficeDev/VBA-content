---
title: TextRange.RotatedBounds Method (PowerPoint)
keywords: vbapp10.chm569036
f1_keywords:
- vbapp10.chm569036
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.RotatedBounds
ms.assetid: 33a4393e-3b87-a4d3-0e16-8230a4beabe3
ms.date: 06/08/2017
---


# TextRange.RotatedBounds Method (PowerPoint)

Returns the coordinates of the vertices of the text bounding box for the specified text range.


## Syntax

 _expression_. **RotatedBounds**( **_X1_**, **_Y1_**, **_X2_**, **_Y2_**, **_X3_**, **_Y3_**, **_X4_**, **_Y4_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _X1, Y1_|Required|**Single**|Returns the position (in points) of the first vertex of the bounding box for the text within the specified text range.|
| _X2, Y2_|Required|**Single**|Returns the position (in points) of the second vertex of the bounding box for the text within the specified text range.|
| _X3, Y3_|Required|**Single**|Returns the position (in points) of the third vertex of the bounding box for the text within the specified text range.|
| _X4, Y4_|Required|**Single**|Returns the position (in points) of the fourth vertex of the bounding box for the text within the specified text range.|

## Example

This example uses the values returned by the arguments of the  **RotatedBounds** method to draw a freeform that has the dimensions of the text bounding box for the third word in the text range in shape one on slide one in the active presentation.


```vb
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim x3 As Single, y3 As Single
Dim x4 As Single, y4 As Single
Dim myDocument As Slide

Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).TextFrame.TextRange.Words(3).RotatedBounds _
    x1, y1, x2, y2, x3, y3, x4, y4

With myDocument.Shapes.BuildFreeform(msoEditingCorner, x1, y1)
    .AddNodes msoSegmentLine, msoEditingAuto, x2, y2
    .AddNodes msoSegmentLine, msoEditingAuto, x3, y3
    .AddNodes msoSegmentLine, msoEditingAuto, x4, y4
    .AddNodes msoSegmentLine, msoEditingAuto, x1, y1
    .ConvertToShape.ZOrder msoSendToBack
End With
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

