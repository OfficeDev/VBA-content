---
title: TextRange.MajorityParagraphFormat Property (Publisher)
keywords: vbapb10.chm5308468
f1_keywords:
- vbapb10.chm5308468
ms.prod: publisher
api_name:
- Publisher.TextRange.MajorityParagraphFormat
ms.assetid: d67e81fe-ab9b-8bfd-c31d-76feb1b6e15b
ms.date: 06/08/2017
---


# TextRange.MajorityParagraphFormat Property (Publisher)

Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object that represents the paragraph formatting applied to most of the paragraphs in a text range.


## Syntax

 _expression_. **MajorityParagraphFormat**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

ParagraphFormat


## Example

This example applies the paragraph formatting applied to a majority of the paragraphs in the first shape to the paragraphs in the second shape on the first page of the active document. This example assumes that there are at least two shapes on page one of the active publication.


```vb
Sub SetFontName() 
 Dim fmt As ParagraphFormat 
 With ActiveDocument.Pages(1) 
 Set fmt = .Shapes(1).TextFrame.TextRange _ 
 .MajorityParagraphFormat 
 .Shapes(2).TextFrame.TextRange.ParagraphFormat = fmt 
 End With 
End Sub
```


