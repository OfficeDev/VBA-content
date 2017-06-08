---
title: ParagraphFormat.SpaceBefore Property (Publisher)
keywords: vbapb10.chm5439497
f1_keywords:
- vbapb10.chm5439497
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SpaceBefore
ms.assetid: ed19a927-67e4-a1b3-06f8-1035c4b0815a
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceBefore Property (Publisher)

Returns or sets a  **Variant** that represents the amount of spacing (in points) before one or more paragraphs. Read/write.


## Syntax

 _expression_. **SpaceBefore**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Variant


## Example

This example sets the spacing before and after the third paragraph in the first shape on the first page of the active publication to 6 points. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub SetSpacingBeforeAfterParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(3).ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
End Sub
```

This example sets spacing before and after all paragraphs in the first shape on the first page of the active publication to 6 points. This example assumes there is at least one shape on the first page of the active publication.




```vb
Sub SetSpacingBeforeAfterAllParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat 
 .SpaceBefore = 12 
 .SpaceAfter = 6 
 End With 
End Sub
```


