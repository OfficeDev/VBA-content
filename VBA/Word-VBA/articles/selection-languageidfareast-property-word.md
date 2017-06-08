---
title: Selection.LanguageIDFarEast Property (Word)
keywords: vbawd10.chm158662810
f1_keywords:
- vbawd10.chm158662810
ms.prod: word
api_name:
- Word.Selection.LanguageIDFarEast
ms.assetid: 59f5b72f-3ba5-cff8-8465-6759d2194d26
ms.date: 06/08/2017
---


# Selection.LanguageIDFarEast Property (Word)

Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .


## Syntax

 _expression_ . **LanguageIDFarEast**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word.


## Example

This example sets the language of the selection to Korean.


```
Selection.LanguageIDFarEast = wdKorean
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

