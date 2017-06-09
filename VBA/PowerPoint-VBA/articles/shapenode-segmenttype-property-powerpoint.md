---
title: ShapeNode.SegmentType Property (PowerPoint)
keywords: vbapp10.chm561004
f1_keywords:
- vbapp10.chm561004
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNode.SegmentType
ms.assetid: 5135d7a7-3ed7-6abd-b072-7456a59aa707
ms.date: 06/08/2017
---


# ShapeNode.SegmentType Property (PowerPoint)

Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only.


## Syntax

 _expression_. **SegmentType**

 _expression_ A variable that represents a **ShapeNode** object.


### Return Value

MsoSegmentType


## Remarks

This property is read-only. Use the  **[SetSegmentType](shapenodes-setsegmenttype-method-powerpoint.md)** method to set the value of this property.

The value returned by the  **SegmentType** property can be one of these **MsoSegmentType** constants. The **SegmentType** property returns **msoSegmentCurve** if the specified node is a control point for a curved segment.


||
|:-----|
|**msoSegmentCurve**|
|**msoSegmentLine**|

## Example

This example changes all straight segments to curved segments in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    n = 1

    While n <= .Count

        If .Item(n).SegmentType = msoSegmentLine Then

            .SetSegmentType n, msoSegmentCurve

        End If

        n = n + 1

    Wend

End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-powerpoint.md)

