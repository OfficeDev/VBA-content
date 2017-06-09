---
title: Paragraph.Space15 Method (Word)
keywords: vbawd10.chm156696890
f1_keywords:
- vbawd10.chm156696890
ms.prod: word
api_name:
- Word.Paragraph.Space15
ms.assetid: c7978808-2a02-609d-1640-b0fef3d24d2a
ms.date: 06/08/2017
---


# Paragraph.Space15 Method (Word)

Formats the specified paragraphs with 1.5-line spacing.


## Syntax

 _expression_ . **Space15**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](paragraph-linespacingrule-property-word.md)** property to set the line spacing for a paragraph. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs(1).Space15 
ActiveDocument.Paragraphs(1).LineSpacingRule = wdLineSpace1pt5
```


## Example

This example changes the first paragraph in the active document to 1.5-line spacing.


```vb
ActiveDocument.Paragraphs(1).Space15
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

