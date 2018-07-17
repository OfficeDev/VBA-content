---
title: Range.LanguageIDOther Property (Word)
keywords: vbawd10.chm157155650
f1_keywords:
- vbawd10.chm157155650
ms.prod: word
api_name:
- Word.Range.LanguageIDOther
ms.assetid: 00b07195-df7d-a979-2534-370cf6540c79
ms.date: 06/08/2017
---


# Range.LanguageIDOther Property (Word)

Returns or sets the language for the specified range. Read/write  **WdLanguageID** .


## Syntax

 _expression_ . **LanguageIDOther**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the language of the selection to French.


```
Selection.Range.LanguageIDOther = wdFrench
```


## See also


#### Concepts


[Range Object](range-object-word.md)

