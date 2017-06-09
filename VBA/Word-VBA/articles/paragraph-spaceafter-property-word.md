---
title: Paragraph.SpaceAfter Property (Word)
keywords: vbawd10.chm156696688
f1_keywords:
- vbawd10.chm156696688
ms.prod: word
api_name:
- Word.Paragraph.SpaceAfter
ms.assetid: 1d720690-f8e3-6b05-f5d2-dd86d29ec4b9
ms.date: 06/08/2017
---


# Paragraph.SpaceAfter Property (Word)

Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceAfter**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets the spacing after the first paragraph in the active document to 12 points.


```vb
ActiveDocument.Paragraphs(1).SpaceAfter = 12
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

