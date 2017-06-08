---
title: Replacement.ParagraphFormat Property (Word)
keywords: vbawd10.chm162594827
f1_keywords:
- vbawd10.chm162594827
ms.prod: word
api_name:
- Word.Replacement.ParagraphFormat
ms.assetid: 0cb9410e-74c9-0fd2-377e-c045dc0274c1
ms.date: 06/08/2017
---


# Replacement.ParagraphFormat Property (Word)

Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified replacement operation. Read/write.


## Syntax

 _expression_ . **ParagraphFormat**

 _expression_ A variable that represents a **[Replacement](replacement-object-word.md)** object.


## Example

This example finds all double-spaced paragraphs in the active document and replaces the formatting with 1.5-line spacing.


```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .ParagraphFormat.Space2 
 .Replacement.ClearFormatting 
 .Replacement.ParagraphFormat.Space15 
 .Execute FindText:="", ReplaceWith:="", _ 
 Replace:=wdReplaceAll 
End With
```


## See also


#### Concepts


[Replacement Object](replacement-object-word.md)

