---
title: ParagraphFormat.SpaceAfter Property (Word)
keywords: vbawd10.chm156434544
f1_keywords:
- vbawd10.chm156434544
ms.prod: word
api_name:
- Word.ParagraphFormat.SpaceAfter
ms.assetid: 804ffaf6-8a74-22a1-7a90-132847b196ee
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceAfter Property (Word)

Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceAfter**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets the spacing after all paragraphs in the active document to 12 points.


```vb
ActiveDocument.Range.ParagraphFormat.SpaceAfter = 12
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

