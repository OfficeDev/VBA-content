---
title: Template.Saved Property (Word)
keywords: vbawd10.chm157941765
f1_keywords:
- vbawd10.chm157941765
ms.prod: word
api_name:
- Word.Template.Saved
ms.assetid: 334069e0-f419-ddf7-0327-6c875bf3b7cd
ms.date: 06/08/2017
---


# Template.Saved Property (Word)

 **True** if the specified template has not changed since it was last saved. **False** if Microsoft Word displays a prompt to save changes when the document is closed. Read/write **Boolean** .


## Syntax

 _expression_ . **Saved**

 _expression_ A variable that represents a **[Template](template-object-word.md)** object.


## Example

This example changes the status of the Normal template to unchanged. If changes were made to the Normal template, the changes are not saved when you exit Word.


```vb
NormalTemplate.Saved = True 
Application.Quit
```


## See also


#### Concepts


[Template Object](template-object-word.md)

