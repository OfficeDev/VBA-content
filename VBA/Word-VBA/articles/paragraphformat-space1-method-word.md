---
title: ParagraphFormat.Space1 Method (Word)
keywords: vbawd10.chm156434745
f1_keywords:
- vbawd10.chm156434745
ms.prod: word
api_name:
- Word.ParagraphFormat.Space1
ms.assetid: 57cc0cea-e50d-affd-1564-30f9240f197b
ms.date: 06/08/2017
---


# ParagraphFormat.Space1 Method (Word)

Single-spaces the specified paragraphs.


## Syntax

 _expression_ . **Space1**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

The exact spacing is determined by the font size of the largest characters in each paragraph.

You can also use the  **[LineSpacingRule](paragraphformat-linespacingrule-property-word.md)** property to set the spacing of paragraphs. The following two statements are equivalent:




```
Selection.ParagraphFormat.Space1 
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
```


## Example

This example changes the first paragraph in the active document to single spacing.


```
Selection.ParagraphFormat.Space1
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

