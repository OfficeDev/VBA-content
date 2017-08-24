---
title: TextRange.InlineShapes Property (Publisher)
keywords: vbapb10.chm5308498
f1_keywords:
- vbapb10.chm5308498
ms.prod: publisher
api_name:
- Publisher.TextRange.InlineShapes
ms.assetid: ffe2d8f2-e1d7-44ea-00fd-3c6523c9fe44
ms.date: 06/08/2017
---


# TextRange.InlineShapes Property (Publisher)

Returns an  **[InlineShapes](inlineshapes-object-publisher.md)** collection, which represents the inline shapes contained within a text range. Read-only.


## Syntax

 _expression_. **InlineShapes**

 _expression_A variable that represents an  **TextRange** object.


### Return Value

InlineShapes


## Remarks

Using  **TextFrame.Story.TextRange.InlineShapes** will return all inline shapes in a text frame, including those that are in overflow. Using **TextFrame.TextRange.InlineShapes** will return only visible inline shapes in a text frame, and not those that are in overflow.


## Example

The following example finds the first shape (a text box) on page one of the active publication. The  **InlineShapes** property is then used to determine whether any inline shapes exist in the text box. If any are found, each inline shape is flipped vertically, and its fore color is set to red.

Note that by using  **TextFrame.Story.TextRange.InlineShapes**, any inline shapes that are in overflow will also be found.




```vb
Dim theShape As Shape 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 For i = 1 To .InlineShapes.Count 
 .InlineShapes(i).Flip (msoFlipVertical) 
 .InlineShapes(i).Fill.ForeColor.RGB = vbRed 
 Next 
 End If 
End With
```


