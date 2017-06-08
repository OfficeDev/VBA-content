---
title: Paragraph.SpaceBefore Property (Word)
keywords: vbawd10.chm156696687
f1_keywords:
- vbawd10.chm156696687
ms.prod: word
api_name:
- Word.Paragraph.SpaceBefore
ms.assetid: 3e9cf50f-5e63-ea24-fe39-7fc9d8690bb4
ms.date: 06/08/2017
---


# Paragraph.SpaceBefore Property (Word)

Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceBefore**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets the spacing before the second paragraph in the active document to 12 points.


```vb
ActiveDocument.Paragraphs(2).SpaceBefore = 12
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

