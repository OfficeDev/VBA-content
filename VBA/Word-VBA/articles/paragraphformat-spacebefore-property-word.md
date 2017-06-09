---
title: ParagraphFormat.SpaceBefore Property (Word)
keywords: vbawd10.chm156434543
f1_keywords:
- vbawd10.chm156434543
ms.prod: word
api_name:
- Word.ParagraphFormat.SpaceBefore
ms.assetid: da20b86e-b69c-f7df-cbaa-46f208ddbdc9
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceBefore Property (Word)

Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceBefore**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets the spacing before all paragraphs in the active document to 12 points.


```vb
ActiveDocument.Range.ParagraphFormat.SpaceBefore = 12
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

