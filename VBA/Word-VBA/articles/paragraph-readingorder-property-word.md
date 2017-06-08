---
title: Paragraph.ReadingOrder Property (Word)
keywords: vbawd10.chm156696779
f1_keywords:
- vbawd10.chm156696779
ms.prod: word
api_name:
- Word.Paragraph.ReadingOrder
ms.assetid: acc70d54-2420-4c03-ab5e-1604f85a6f66
ms.date: 06/08/2017
---


# Paragraph.ReadingOrder Property (Word)

Returns or sets the reading order of the specified paragraph without changing the alignment. Read/write  **WdReadingOrder** .


## Syntax

 _expression_ . **ReadingOrder**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

Use the  **[LtrPara](selection-ltrpara-method-word.md)** , **[LtrRun](selection-ltrrun-method-word.md)** , **[RtlPara](selection-rtlpara-method-word.md)** , and **[RtlRun](selection-rtlrun-method-word.md)** methods of the **[Selection](selection-object-word.md)** object to change the paragraph alignment along with the reading order.


## Example

This example sets the reading order of the first paragraph to right-to-left.


```vb
ActiveDocument.Paragraphs(1).ReadingOrder = _ 
 wdReadingOrderRtl
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

