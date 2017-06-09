---
title: Paragraphs.OpenUp Method (Word)
keywords: vbawd10.chm156762414
f1_keywords:
- vbawd10.chm156762414
ms.prod: word
api_name:
- Word.Paragraphs.OpenUp
ms.assetid: 0998519f-5fdc-3ac1-488f-03ff179be1c9
ms.date: 06/08/2017
---


# Paragraphs.OpenUp Method (Word)

Sets spacing before the specified paragraphs to 12 points.


## Syntax

 _expression_ . **OpenUp**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

You can also use the  **[SpaceBefore](paragraphs-spacebefore-property-word.md)** property to set the spacing before paragraphs. The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs.OpenUp 
ActiveDocument.Paragraphs.SpaceBefore = 12
```


## Example

This example changes the formatting of the second paragraph in the active document to leave 12 points of space before the paragraph.


```vb
ActiveDocument.Paragraphs(2).OpenUp
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

