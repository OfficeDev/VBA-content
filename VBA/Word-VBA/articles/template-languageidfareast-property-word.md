---
title: Template.LanguageIDFarEast Property (Word)
keywords: vbawd10.chm157941771
f1_keywords:
- vbawd10.chm157941771
ms.prod: word
api_name:
- Word.Template.LanguageIDFarEast
ms.assetid: d9798c5a-1362-a713-0357-2599d5038c18
ms.date: 06/08/2017
---


# Template.LanguageIDFarEast Property (Word)

Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .


## Syntax

 _expression_ . **LanguageIDFarEast**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


## Remarks

This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word.


## Example

This example sets the language of the selection to Korean.


```
NormalTemplate.LanguageIDFarEast = wdKorean
```


## See also


#### Concepts


[Template Object](template-object-word.md)

