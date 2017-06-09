---
title: Template.KerningByAlgorithm Property (Word)
keywords: vbawd10.chm157941772
f1_keywords:
- vbawd10.chm157941772
ms.prod: word
api_name:
- Word.Template.KerningByAlgorithm
ms.assetid: 4812a92c-8886-6c52-4b26-6fc50e270f21
ms.date: 06/08/2017
---


# Template.KerningByAlgorithm Property (Word)

 **True** if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write **Boolean** .


## Syntax

 _expression_ . **KerningByAlgorithm**

 _expression_ A variable that represents a **[Template](template-object-word.md)** object.


## Example

This example sets Microsoft Word to kern half-width Latin characters and punctuation marks in the normal template.


```vb
NormalTemplate.KerningByAlgorithm = True
```


## See also


#### Concepts


[Template Object](template-object-word.md)

