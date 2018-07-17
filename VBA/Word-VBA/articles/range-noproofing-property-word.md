---
title: Range.NoProofing Property (Word)
keywords: vbawd10.chm157155651
f1_keywords:
- vbawd10.chm157155651
ms.prod: word
api_name:
- Word.Range.NoProofing
ms.assetid: 0344239d-10bc-0e3e-9601-41c3c3bb6227
ms.date: 06/08/2017
---


# Range.NoProofing Property (Word)

 **True** if the spelling and grammar checker ignores the specified text. Read/write **Long** .


## Syntax

 _expression_ . **NoProofing**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if the **NoProofing** property is set to **True** for only some of the specified text.


## Example

This example marks the current selection to be ignored by the spelling and grammar checker.


```vb
Selection.Range.NoProofing = True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

