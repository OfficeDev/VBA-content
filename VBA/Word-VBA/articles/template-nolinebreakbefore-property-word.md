---
title: Template.NoLineBreakBefore Property (Word)
keywords: vbawd10.chm157941775
f1_keywords:
- vbawd10.chm157941775
ms.prod: word
api_name:
- Word.Template.NoLineBreakBefore
ms.assetid: 47a827aa-a436-e1c5-1d32-748eb2c833df
ms.date: 06/08/2017
---


# Template.NoLineBreakBefore Property (Word)

Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write  **String** .


## Syntax

 _expression_ . **NoLineBreakBefore**

 _expression_ A variable that represents a **[Template](template-object-word.md)** object.


## Example

This example sets "!", ")", and "]" as the kinsoku characters before which Word will not break a line in the active document.


```
NormalTemplate.NoLineBreakBefore = "!)]"
```


## See also


#### Concepts


[Template Object](template-object-word.md)

