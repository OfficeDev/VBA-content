---
title: ShapeRange.InlineTextRange Property (Publisher)
keywords: vbapb10.chm2294023
f1_keywords:
- vbapb10.chm2294023
ms.prod: publisher
api_name:
- Publisher.ShapeRange.InlineTextRange
ms.assetid: 5d7f3dfa-3e23-85c6-50cf-a6f960ccabfc
ms.date: 06/08/2017
---


# ShapeRange.InlineTextRange Property (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that reflects the position of the inline shape in its containing text range. Read-only.


## Syntax

 _expression_. **InlineTextRange**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The returned text range will contain a single object representing the inline shape. An automation error is returned if the shape is not inline.


## Example

The following example finds the first shape (a text box) on the first page of the publication, and determines if the text range within the text box contains inline shapes. If inline shapes are found, the  **InlineTextRange** property is used to represent the inline shape after a block of text is inserted.


```vb
Dim theShape As Shape 
Dim theTextRange As TextRange 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
If Not theShape.IsInline = True Then 
 With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 Set theTextRange = theShape.TextFrame.Story.TextRange 
 For i = 1 To .InlineShapes.Count 
 With .InlineShapes(i) 
 .InlineTextRange.InsertAfter (" (Figure " &; i &; ") ") 
 End With 
 Next 
 End If 
 End With 
End If
```


