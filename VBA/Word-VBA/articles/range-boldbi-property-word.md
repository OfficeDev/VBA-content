---
title: Range.BoldBi Property (Word)
keywords: vbawd10.chm157155728
f1_keywords:
- vbawd10.chm157155728
ms.prod: word
api_name:
- Word.Range.BoldBi
ms.assetid: 80a4e893-0337-41ef-5a45-506deea43f29
ms.date: 06/08/2017
---


# Range.BoldBi Property (Word)

 **True** if the font or range is formatted as bold. Returns **True** , **False** , or **wdUndefined** (for a mixture of bold and non-bold text). Can be set to **True** , **False** , or **wdToggle** . Read/write **Long** .


## Syntax

 _expression_ . **BoldBi**

 _expression_ An expression that returns a **[Range](range-object-word.md)** object.


## Example

This example makes the first paragraph in the active right-to-left language document bold.


```vb
ActiveDocument.Paragraphs(1).Range.BoldBi = True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

