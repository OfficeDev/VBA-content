---
title: ParagraphFormat.OpenOrCloseUp Method (Word)
keywords: vbawd10.chm156434735
f1_keywords:
- vbawd10.chm156434735
ms.prod: word
api_name:
- Word.ParagraphFormat.OpenOrCloseUp
ms.assetid: 7cf08077-e3e5-4886-e88f-fd12c2961058
ms.date: 06/08/2017
---


# ParagraphFormat.OpenOrCloseUp Method (Word)

Toggles the spacing before the specified paragraphs.


## Syntax

 _expression_ . **OpenOrCloseUp**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

If the spacing before the specified paragraphs is 0 (zero), this method sets spacing to 12 points. If the spacing before the paragraphs is greater than 0 (zero), this method sets spacing to 0 (zero).


## Example

This example toggles the formatting of the first paragraph in the active document to either add 12 points of space before the paragraph or leave no space before it.


```
Selection.ParagraphFormat.OpenOrCloseUp
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

