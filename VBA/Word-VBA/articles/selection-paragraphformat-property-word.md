---
title: Selection.ParagraphFormat Property (Word)
keywords: vbawd10.chm158663758
f1_keywords:
- vbawd10.chm158663758
ms.prod: word
api_name:
- Word.Selection.ParagraphFormat
ms.assetid: 3a3a3b4e-396f-fbe5-dc30-649ef7a9a8f9
ms.date: 06/08/2017
---


# Selection.ParagraphFormat Property (Word)

Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified selection. Read/write.


## Syntax

 _expression_ . **ParagraphFormat**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example sets the paragraph formatting for the current selection to be right-aligned.


```
Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

