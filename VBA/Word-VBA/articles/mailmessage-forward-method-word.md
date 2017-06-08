---
title: MailMessage.Forward Method (Word)
keywords: vbawd10.chm163184979
f1_keywords:
- vbawd10.chm163184979
ms.prod: word
api_name:
- Word.MailMessage.Forward
ms.assetid: 3ae7a3bc-9cc1-82eb-eff5-ea4a99fe181f
ms.date: 06/08/2017
---


# MailMessage.Forward Method (Word)

Opens a new e-mail message with an empty  **To** line for forwarding the active message.


## Syntax

 _expression_ . **Forward**

 _expression_ Required. A variable that represents a **[MailMessage](mailmessage-object-word.md)** object.


## Remarks

This method is available only if you are using Word as your e-mail editor.


## Example

This example opens a new e-mail message for forwarding the active message.


```vb
Application.MailMessage.Forward
```


## See also


#### Concepts


[MailMessage Object](mailmessage-object-word.md)

