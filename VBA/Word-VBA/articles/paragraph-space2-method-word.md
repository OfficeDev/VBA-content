---
title: Paragraph.Space2 Method (Word)
keywords: vbawd10.chm156696891
f1_keywords:
- vbawd10.chm156696891
ms.prod: word
api_name:
- Word.Paragraph.Space2
ms.assetid: 51feb546-a6e4-4f8c-74b8-a6cf7b9c068c
ms.date: 06/08/2017
---


# Paragraph.Space2 Method (Word)

Double-spaces the specified paragraphs.


## Syntax

 _expression_ . **Space2**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](paragraph-linespacingrule-property-word.md)** property to set the line spacing for a paragraph. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs(1).Space2 
ActiveDocument.Paragraphs(1).LineSpacingRule = wdLineSpaceDouble
```


## Example

This example changes the first paragraph in the selection to double spacing.


```
Selection.Paragraphs(1).Space2
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

