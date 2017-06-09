---
title: Paragraph.FarEastLineBreakControl Property (Word)
keywords: vbawd10.chm156696693
f1_keywords:
- vbawd10.chm156696693
ms.prod: word
api_name:
- Word.Paragraph.FarEastLineBreakControl
ms.assetid: 974b326b-5acc-bafd-6b0a-b9e6657d0058
ms.date: 06/08/2017
---


# Paragraph.FarEastLineBreakControl Property (Word)

 **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **FarEastLineBreakControl**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets Word to apply East Asian line-breaking rules to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).FarEastLineBreakControl = True
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

