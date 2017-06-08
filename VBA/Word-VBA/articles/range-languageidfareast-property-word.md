---
title: Range.LanguageIDFarEast Property (Word)
keywords: vbawd10.chm157155649
f1_keywords:
- vbawd10.chm157155649
ms.prod: word
api_name:
- Word.Range.LanguageIDFarEast
ms.assetid: 324eaba2-2a48-71e3-6a96-9b7a092d0c6d
ms.date: 06/08/2017
---


# Range.LanguageIDFarEast Property (Word)

Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .


## Syntax

 _expression_ . **LanguageIDFarEast**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the language of the first paragraph in the active document to Korean.


```vb
ActiveDocument.Paragraphs(1).Range.LanguageIDFarEast = wdKorean
```


## See also


#### Concepts


[Range Object](range-object-word.md)

