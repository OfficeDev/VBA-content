---
title: Paragraph.Alignment Property (Word)
keywords: vbawd10.chm156696677
f1_keywords:
- vbawd10.chm156696677
ms.prod: word
api_name:
- Word.Paragraph.Alignment
ms.assetid: 0142adc2-624c-eb9b-7eca-b24a2f16573f
ms.date: 06/08/2017
---


# Paragraph.Alignment Property (Word)

Returns or sets a  **WdParagraphAlignment** constant that represents the alignment for the specified paragraphs. Read/write.


## Syntax

 _expression_ . **Alignment**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

Some of the  **WdParagraphAlignment** constants, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example right-aligns the first paragraph in the active document.


```vb
Sub AlignParagraph() 
 ActiveDocument.Paragraphs(1).Alignment = _ 
 wdAlignParagraphRight 
End Sub
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

