---
title: ParagraphFormat.SpaceAfter Property (Publisher)
keywords: vbapb10.chm5439496
f1_keywords:
- vbapb10.chm5439496
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SpaceAfter
ms.assetid: 52f65636-862d-442e-e66f-5ff5c79ee7b0
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceAfter Property (Publisher)

Returns or sets a  **Variant** that represents the amount of spacing (in points) after one or more paragraphs. Read/write.


## Syntax

 _expression_. **SpaceAfter**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Variant


## Example

This example sets the spacing before and after the third paragraph in the first shape on the first page of the active publication to 6 points.


```vb
Sub SetSpacingBeforeAfterParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(3).ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
End Sub
```

This example sets spacing before and after all paragraphs in the first shape on the first page of the active publication to 6 points.




```vb
Sub SetSpacingBeforeAfterAllParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat 
 .SpaceBefore = 12 
 .SpaceAfter = 6 
 End With 
End Sub
```


