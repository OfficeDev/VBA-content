---
title: ParagraphFormat.ReadingOrder Property (Word)
keywords: vbawd10.chm156434563
f1_keywords:
- vbawd10.chm156434563
ms.prod: word
api_name:
- Word.ParagraphFormat.ReadingOrder
ms.assetid: 4a22e638-2af8-096a-d45c-2eed21dc8002
ms.date: 06/08/2017
---


# ParagraphFormat.ReadingOrder Property (Word)

Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder** .


## Syntax

 _expression_ . **ReadingOrder**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

Use the  **[LtrPara](selection-ltrpara-method-word.md)** , **[LtrRun](selection-ltrrun-method-word.md)** , **[RtlPara](selection-rtlpara-method-word.md)** , and **[RtlRun](selection-rtlrun-method-word.md)** methods of the **Selection** object to change the paragraph alignment along with the reading order.


## Example

This example sets the reading order of the first paragraph to right-to-left.


```vb
ActiveDocument.Paragraphs(1).ReadingOrder = _ 
 wdReadingOrderRtl
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

