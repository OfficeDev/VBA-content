---
title: Application.EmailOptions Property (Word)
keywords: vbawd10.chm158335365
f1_keywords:
- vbawd10.chm158335365
ms.prod: word
api_name:
- Word.Application.EmailOptions
ms.assetid: 28547346-6119-b763-339e-b04af1c8268f
ms.date: 06/08/2017
---


# Application.EmailOptions Property (Word)

Returns an  **[EmailOptions](emailoptions-object-word.md)** object that represents the global preferences for e-mail authoring. Read-only.


## Syntax

 _expression_ . **EmailOptions**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example sets Microsoft Word to mark comments in e-mail messages.


```vb
Application.EmailOptions.MarkComments = True
```


## See also


#### Concepts


[Application Object](application-object-word.md)

