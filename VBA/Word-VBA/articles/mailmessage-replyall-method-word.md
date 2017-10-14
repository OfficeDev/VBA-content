---
title: MailMessage.ReplyAll Method (Word)
keywords: vbawd10.chm163184983
f1_keywords:
- vbawd10.chm163184983
ms.prod: word
api_name:
- Word.MailMessage.ReplyAll
ms.assetid: cc7aa537-573f-f2b2-14a1-3443ed622f56
ms.date: 06/08/2017
---


# MailMessage.ReplyAll Method (Word)

Opens a new e-mail message — with the sender's and all other recipients' addresses on the  **To** and **Cc** lines, as appropriate — for replying to the active message.


## Syntax

 _expression_ . **ReplyAll**

 _expression_ Required. A variable that represents a **[MailMessage](mailmessage-object-word.md)** object.


## Example

This example opens a new e-mail message for replying to the active message.


```vb
Application.MailMessage.ReplyAll
```


## See also


#### Concepts


[MailMessage Object](mailmessage-object-word.md)

