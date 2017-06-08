---
title: ParagraphFormat.WordWrap Property (Word)
keywords: vbawd10.chm156434550
f1_keywords:
- vbawd10.chm156434550
ms.prod: word
api_name:
- Word.ParagraphFormat.WordWrap
ms.assetid: da5e67c3-405d-8adb-5cec-321464030f08
ms.date: 06/08/2017
---


# ParagraphFormat.WordWrap Property (Word)

 **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. Read/write **Long** .


## Syntax

 _expression_ . **WordWrap**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if it is set to **True** for only some of the specified paragraphs or text frames. This usage may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example sets Microsoft Word to wrap Latin text in the middle of a word in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).WordWrap = True
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

