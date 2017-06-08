---
title: Application.MailMessage Property (Word)
keywords: vbawd10.chm158335324
f1_keywords:
- vbawd10.chm158335324
ms.prod: word
api_name:
- Word.Application.MailMessage
ms.assetid: 82bca039-0b6b-4489-27bf-18746dc639d2
ms.date: 06/08/2017
---


# Application.MailMessage Property (Word)

Returns a  **[MailMessage](mailmessage-object-word.md)** object that represents the active e-mail message.


## Syntax

 _expression_ . **MailMessage**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example displays the  **Select Names** dialog box for the active e-mail message.


```vb
Application.MailMessage.DisplaySelectNamesDialog
```


## See also


#### Concepts


[Application Object](application-object-word.md)

