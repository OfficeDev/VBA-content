---
title: Paragraphs.Space1 Method (Word)
keywords: vbawd10.chm156762425
f1_keywords:
- vbawd10.chm156762425
ms.prod: word
api_name:
- Word.Paragraphs.Space1
ms.assetid: fe426595-427a-51bd-3e65-48d3b3e4c78d
ms.date: 06/08/2017
---


# Paragraphs.Space1 Method (Word)

Single-spaces the specified paragraphs.


## Syntax

 _expression_ . **Space1**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

The exact spacing is determined by the font size of the largest characters in each paragraph.

You can also use the  **[LineSpacingRule](paragraphs-linespacingrule-property-word.md)** property to set paragraph spacing. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs.Space1 
ActiveDocument.Paragraphs.LineSpacingRule = wdLineSpaceSingle
```


## Example

This example changes all paragraphs in the active document to single spacing.


```vb
ActiveDocument.Paragraphs.Space1
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

