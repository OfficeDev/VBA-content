---
title: ParagraphFormat.FarEastLineBreakControl Property (Word)
keywords: vbawd10.chm156434549
f1_keywords:
- vbawd10.chm156434549
ms.prod: word
api_name:
- Word.ParagraphFormat.FarEastLineBreakControl
ms.assetid: 554a0097-5402-2b40-face-c9ec942ad3e1
ms.date: 06/08/2017
---


# ParagraphFormat.FarEastLineBreakControl Property (Word)

 **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **FarEastLineBreakControl**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets Word to apply East Asian line-breaking rules to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).FarEastLineBreakControl = True
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

