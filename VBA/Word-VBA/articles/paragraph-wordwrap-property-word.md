---
title: Paragraph.WordWrap Property (Word)
keywords: vbawd10.chm156696694
f1_keywords:
- vbawd10.chm156696694
ms.prod: word
api_name:
- Word.Paragraph.WordWrap
ms.assetid: d7e4da55-8ef8-55f5-ad4d-8dc487b737ce
ms.date: 06/08/2017
---


# Paragraph.WordWrap Property (Word)

 **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. Read/write **Long** .


## Syntax

 _expression_ . **WordWrap**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if it's set to **True** for only some of the specified paragraphs or text frames. This usage may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example sets Microsoft Word to wrap Latin text in the middle of a word in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).WordWrap = False
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

