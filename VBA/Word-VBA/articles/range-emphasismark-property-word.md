---
title: Range.EmphasisMark Property (Word)
keywords: vbawd10.chm157155468
f1_keywords:
- vbawd10.chm157155468
ms.prod: WORD
api_name:
- Word.Range.EmphasisMark
ms.assetid: 6f0f7d19-efba-8fee-7e6c-abb1defe8529
---


# Range.EmphasisMark Property (Word)

Returns or sets the emphasis mark for a character or designated character string. Read/write  **WdEmphasisMark** .


## Syntax

 _expression_ . **EmphasisMark**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the emphasis mark over the fourth word in the active document to a comma.


```vb
ActiveDocument.Words(4).EmphasisMark = wdEmphasisMarkOverComma
```


## See also


#### Concepts


[Range Object](range-object-word.md)

