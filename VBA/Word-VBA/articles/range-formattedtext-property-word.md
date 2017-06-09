---
title: Range.FormattedText Property (Word)
keywords: vbawd10.chm157155330
f1_keywords:
- vbawd10.chm157155330
ms.prod: word
api_name:
- Word.Range.FormattedText
ms.assetid: 26221da8-e3d7-4da5-f23a-cd678d8ab2f5
ms.date: 06/08/2017
---


# Range.FormattedText Property (Word)

Returns or sets a  **Range** object that includes the formatted text in the specified range or selection. Read/write.


## Syntax

 _expression_ . **FormattedText**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

This property returns a  **Range** object with the character formatting and text from the specified range or selection. Paragraph formatting is included in the **Range** object if there is a paragraph mark in the range or selection.

When you set this property, the text in the range is replaced with formatted text. If you do not want to replace the existing text, use the  **Collapse** method before using this property (see the first example).


## See also


#### Concepts


[Range Object](range-object-word.md)

