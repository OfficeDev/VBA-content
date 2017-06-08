---
title: Template.NoLineBreakAfter Property (Word)
keywords: vbawd10.chm157941776
f1_keywords:
- vbawd10.chm157941776
ms.prod: word
api_name:
- Word.Template.NoLineBreakAfter
ms.assetid: efe38080-35b3-16d4-ce5c-0acb9a2a52ad
ms.date: 06/08/2017
---


# Template.NoLineBreakAfter Property (Word)

Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write  **String** .


## Syntax

 _expression_ . **NoLineBreakAfter**

 _expression_ A variable that represents a **[Template](template-object-word.md)** object.


## Example

This example sets "$", "(", "[", "\", and "{" as the kinsoku characters after which Microsoft Word will not break a line in the active document.


```vb
ActiveDocument.NoLineBreakAfter = "$([\{"
```


## See also


#### Concepts


[Template Object](template-object-word.md)

