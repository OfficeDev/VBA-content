---
title: ParagraphFormat.Space2 Method (Word)
keywords: vbawd10.chm156434747
f1_keywords:
- vbawd10.chm156434747
ms.prod: word
api_name:
- Word.ParagraphFormat.Space2
ms.assetid: 7173f5b8-961b-e93f-e4b6-fedad6da8d1d
ms.date: 06/08/2017
---


# ParagraphFormat.Space2 Method (Word)

Double-spaces the specified paragraphs.


## Syntax

 _expression_ . **Space2**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](paragraphformat-linespacingrule-property-word.md)** property to set the spacing of paragraphs. For example, the following two statements are equivalent:




```
Selection.ParagraphFormat.Space2 
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble
```


## Example

This example changes the first paragraph in the selection to double spacing.


```
Selection.ParagraphFormat.Space2
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

