---
title: Paragraphs.Space2 Method (Word)
keywords: vbawd10.chm156762427
f1_keywords:
- vbawd10.chm156762427
ms.prod: word
api_name:
- Word.Paragraphs.Space2
ms.assetid: dfd70842-8a1b-8266-7c37-1b8d61c046ae
ms.date: 06/08/2017
---


# Paragraphs.Space2 Method (Word)

Double-spaces the specified paragraphs. .


## Syntax

 _expression_ . **Space2**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](paragraphs-linespacingrule-property-word.md)** property to set paragraph spacing. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs.Space2 
ActiveDocument.Paragraphs.LineSpacingRule = wdLineSpaceDouble
```


## Example

This example changes all selected paragraphs to double spacing.


```
Selection.Paragraphs.Space2
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

