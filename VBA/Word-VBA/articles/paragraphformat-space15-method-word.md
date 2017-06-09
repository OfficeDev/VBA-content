---
title: ParagraphFormat.Space15 Method (Word)
keywords: vbawd10.chm156434746
f1_keywords:
- vbawd10.chm156434746
ms.prod: word
api_name:
- Word.ParagraphFormat.Space15
ms.assetid: 6621d8e8-c207-0862-ddd4-33cb5bcd9cbc
ms.date: 06/08/2017
---


# ParagraphFormat.Space15 Method (Word)

Formats the specified paragraphs with 1.5-line spacing.


## Syntax

 _expression_ . **Space15**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](paragraphformat-linespacingrule-property-word.md)** property to set the spacing of paragraphs. The following two statements are equivalent:




```
Selection.ParagraphFormat.Space15 
Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
```


## Example

This example changes the first paragraph in the active document to 1.5-line spacing.


```
Selection.ParagraphFormat.Space15
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

