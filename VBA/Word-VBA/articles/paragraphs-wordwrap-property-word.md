---
title: Paragraphs.WordWrap Property (Word)
keywords: vbawd10.chm156762230
f1_keywords:
- vbawd10.chm156762230
ms.prod: word
api_name:
- Word.Paragraphs.WordWrap
ms.assetid: bf77cc49-c440-3c8e-7384-721658207386
ms.date: 06/08/2017
---


# Paragraphs.WordWrap Property (Word)

 **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **WordWrap**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if it's set to **True** for only some of the specified paragraphs. This property may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example sets Microsoft Word to wrap Latin text in the middle of a word in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).WordWrap = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

