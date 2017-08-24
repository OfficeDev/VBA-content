---
title: Shape.IsInline Property (Publisher)
keywords: vbapb10.chm5308692
f1_keywords:
- vbapb10.chm5308692
ms.prod: publisher
api_name:
- Publisher.Shape.IsInline
ms.assetid: 5c5c6181-070f-2a66-8d70-2d6372cb365e
ms.date: 06/08/2017
---


# Shape.IsInline Property (Publisher)

Returns an  **MsoTriState** constant that specifies whether a shape is inline (contained in a text run). Read-only.


## Syntax

 _expression_. **IsInline**

 _expression_A variable that represents a  **Shape** object.


## Example

This example tests the first shape (a text frame) on the first page of the publication to see if it is inline. If it is not, a search is done within that shape to find any inline shapes within the text frame. Any inline shapes that are found have the  **ForeColor** property set to red.


```vb
Dim theShape As Shape 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
If Not theShape.IsInline = True Then 
 With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 For i = 1 To .InlineShapes.Count 
 .InlineShapes(i).Select 
 .InlineShapes(i).Fill.ForeColor.RGB = vbRed 
 Next 
 End If 
 End With 
End If
```


