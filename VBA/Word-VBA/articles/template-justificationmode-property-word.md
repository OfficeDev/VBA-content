---
title: Template.JustificationMode Property (Word)
keywords: vbawd10.chm157941773
f1_keywords:
- vbawd10.chm157941773
ms.prod: word
api_name:
- Word.Template.JustificationMode
ms.assetid: 914994e8-8ea3-4119-271c-193970da060c
ms.date: 06/08/2017
---


# Template.JustificationMode Property (Word)

Returns or sets the character spacing adjustment for the specified template. Read/write  **[WdJustificationMode](wdjustificationmode-enumeration-word.md)** .


## Syntax

 _expression_ . **JustificationMode**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


## Example

This example sets Microsoft Word to compress only punctuation marks when adjusting character spacing.


```
NormalTemplate.JustificationMode = wdJustificationModeCompressKana
```


## See also


#### Concepts


[Template Object](template-object-word.md)

