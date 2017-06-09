---
title: ParagraphFormat.OpenUp Method (Word)
keywords: vbawd10.chm156434734
f1_keywords:
- vbawd10.chm156434734
ms.prod: word
api_name:
- Word.ParagraphFormat.OpenUp
ms.assetid: 1473b383-816f-087a-073a-5afc5f530c3a
ms.date: 06/08/2017
---


# ParagraphFormat.OpenUp Method (Word)

Sets spacing before the specified paragraphs to 12 points.


## Syntax

 _expression_ . **OpenUp**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

You can also use the  **[SpaceBefore](paragraphformat-spacebefore-property-word.md)** property to set the spacing of paragraphs. The following two statements are equivalent:


```
Selection.ParagraphFormat.OpenUp 
Selection.ParagraphFormat.SpaceBefore = 12
```


## Example

This example changes the formatting of the second paragraph in the active document to leave 12 points of space before the paragraph.


```
Selection.ParagraphFormat.OpenUp
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

