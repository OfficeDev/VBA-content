---
title: ShapeNode.SegmentType Property (Word)
keywords: vbawd10.chm164429926
f1_keywords:
- vbawd10.chm164429926
ms.prod: word
api_name:
- Word.ShapeNode.SegmentType
ms.assetid: d6872a73-6021-8a93-5b1f-95e3349cc818
ms.date: 06/08/2017
---


# ShapeNode.SegmentType Property (Word)

Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only  **MsoSegmentType** .


## Syntax

 _expression_ . **SegmentType**

 _expression_ Required. A variable that represents a **[ShapeNode](shapenode-object-word.md)** object.


## Remarks

If the specified node is a control point for a curved segment, this property returns  **msoSegmentCurve** .

Use the  **SetSegmentType** method to set the value of this property.


## Example

This example changes all straight segments to curved segments in shape three on myDocument. Shape three must be a freeform drawing.


```vb
Set myDocument = ActiveDocument 
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


[ShapeNode Object](shapenode-object-word.md)

