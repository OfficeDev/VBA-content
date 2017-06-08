---
title: MailMessage.DisplayMoveDialog Method (Word)
keywords: vbawd10.chm163184976
f1_keywords:
- vbawd10.chm163184976
ms.prod: word
api_name:
- Word.MailMessage.DisplayMoveDialog
ms.assetid: e913a4f3-e970-ae2f-84b1-c239cc57a15f
ms.date: 06/08/2017
---


# MailMessage.DisplayMoveDialog Method (Word)

Displays the  **Move** dialog box, in which the user can specify a new location for the active e-mail message in an available message store.


## Syntax

 _expression_ . **DisplayMoveDialog**

 _expression_ Required. A variable that represents a **[MailMessage](mailmessage-object-word.md)** object.


## Remarks

This method is available only if you are using Word as your e-mail editor.


## Example

This example displays the  **Move** dialog box for the active e-mail message.


```vb
Application.MailMessage.DisplayMoveDialog
```


## See also


#### Concepts


[MailMessage Object](mailmessage-object-word.md)

