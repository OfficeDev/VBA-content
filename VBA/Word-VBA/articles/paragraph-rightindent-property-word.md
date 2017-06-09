---
title: Paragraph.RightIndent Property (Word)
keywords: vbawd10.chm156696682
f1_keywords:
- vbawd10.chm156696682
ms.prod: word
api_name:
- Word.Paragraph.RightIndent
ms.assetid: 238c2942-14a9-4295-49ed-d4ada5aebd0f
ms.date: 06/08/2017
---


# Paragraph.RightIndent Property (Word)

Returns or sets the right indent (in points) for the specified paragraph. Read/write  **Single** .


## Syntax

 _expression_ . **RightIndent**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets the right indent for the first paragraph in the active document to 1 inch from the right margin. The  **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs(1).RightIndent = InchesToPoints(1)
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

