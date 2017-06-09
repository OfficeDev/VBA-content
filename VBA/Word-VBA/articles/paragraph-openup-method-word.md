---
title: Paragraph.OpenUp Method (Word)
keywords: vbawd10.chm156696878
f1_keywords:
- vbawd10.chm156696878
ms.prod: word
api_name:
- Word.Paragraph.OpenUp
ms.assetid: 660d5595-cf12-db3d-e4d2-0d4880d3df7a
ms.date: 06/08/2017
---


# Paragraph.OpenUp Method (Word)

Sets spacing before the specified paragraphs to 12 points.


## Syntax

 _expression_ . **OpenUp**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

You can also use the  **[SpaceBefore](paragraph-spacebefore-property-word.md)** property to set the spacing for a paragraph. The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs(1).OpenUp 
ActiveDocument.Paragraphs(1).SpaceBefore = 12
```


## Example

This example changes the formatting of the second paragraph in the active document to leave 12 points of space before the paragraph.


```vb
ActiveDocument.Paragraphs(2).OpenUp
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

