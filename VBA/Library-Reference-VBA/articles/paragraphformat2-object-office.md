---
title: ParagraphFormat2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ParagraphFormat2
ms.assetid: 05ff2b24-9603-f923-d053-e736fb2ba389
---


# ParagraphFormat2 Object (Office)

Represents the paragraph formatting of a text range.


## Example

The following example left aligns the paragraphs in shape two on slide one in the active PowerPoint presentation.


```vb
ActivePresentation.Slides(1).Shapes(2).TextFrame2.TextRange2 _ 
 .ParagraphFormat2.Alignment = ppAlignLeft 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

