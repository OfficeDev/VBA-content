---
title: Range.ParagraphFormat Property (Word)
keywords: vbawd10.chm157156430
f1_keywords:
- vbawd10.chm157156430
ms.prod: word
api_name:
- Word.Range.ParagraphFormat
ms.assetid: 98afe866-4d92-7a1d-f5c6-a0128d247df0
ms.date: 06/08/2017
---


# Range.ParagraphFormat Property (Word)

Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified range. Read/write.


## Syntax

 _expression_ . **ParagraphFormat**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets paragraph formatting for a range that includes the entire contents of MyDoc.doc. Paragraphs in this document are double-spaced and have a custom tab stop at 0.25 inch.


```vb
Set myRange = Documents("MyDoc.doc").Content 
With myRange.ParagraphFormat 
 .Space2 
 .TabStops.Add Position:=InchesToPoints(.25) 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

