---
title: Email.CurrentEmailAuthor Property (Word)
keywords: vbawd10.chm165478505
f1_keywords:
- vbawd10.chm165478505
ms.prod: word
api_name:
- Word.Email.CurrentEmailAuthor
ms.assetid: a317b265-f712-2954-aeaf-3a17da38d380
ms.date: 06/08/2017
---


# Email.CurrentEmailAuthor Property (Word)

Returns an  **[EmailAuthor](emailauthor-object-word.md)** object that represents the author of the current e-mail message. Read-only.


## Syntax

 _expression_ . **CurrentEmailAuthor**

 _expression_ A variable that represents a **[Email](email-object-word.md)** object.


## Example

This example returns the name of the style associated with the current e-mail author.


```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```


## See also


#### Concepts


[Email Object](email-object-word.md)

