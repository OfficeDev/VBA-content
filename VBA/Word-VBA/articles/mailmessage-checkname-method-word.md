---
title: MailMessage.CheckName Method (Word)
keywords: vbawd10.chm163184974
f1_keywords:
- vbawd10.chm163184974
ms.prod: word
api_name:
- Word.MailMessage.CheckName
ms.assetid: 2888dfb7-5773-cbf8-8865-c90875411476
ms.date: 06/08/2017
---


# MailMessage.CheckName Method (Word)

Validates the e-mail addresses that appear in the  **To**,  **Cc**, and  **Bcc** lines in the active e-mail message.


## Syntax

 _expression_ . **CheckName**

 _expression_ Required. A variable that represents a **[MailMessage](mailmessage-object-word.md)** object.


## Remarks

This method is available only if you are using Word as your e-mail editor. If the names cannot be validated, the  **Check Names** dialog box is displayed.


## Example

This example validates the e-mail addresses that appear in the active e-mail message.


```vb
Application.MailMessage.CheckName
```


## See also


#### Concepts


[MailMessage Object](mailmessage-object-word.md)

