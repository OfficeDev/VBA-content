---
title: Paragraphs.ReadingOrder Property (Word)
keywords: vbawd10.chm156762243
f1_keywords:
- vbawd10.chm156762243
ms.prod: word
api_name:
- Word.Paragraphs.ReadingOrder
ms.assetid: 9f3fccf3-7474-231d-21c7-f719174d7c82
ms.date: 06/08/2017
---


# Paragraphs.ReadingOrder Property (Word)

Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder** .


## Syntax

 _expression_ . **ReadingOrder**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

Use the  **[LtrPara](selection-ltrpara-method-word.md)** , **[LtrRun](selection-ltrrun-method-word.md)** , **[RtlPara](selection-rtlpara-method-word.md)** , and **[RtlRun](selection-rtlrun-method-word.md)** methods of the **[Selection](selection-object-word.md)** object to change the paragraph alignment along with the reading order.


## Example

This example sets the reading order of all paragraphs in the active document to right-to-left.


```vb
ActiveDocument.Paragraphs.ReadingOrder = _ 
 wdReadingOrderRtl
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

