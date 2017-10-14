---
title: MailMessage.DisplayProperties Method (Word)
keywords: vbawd10.chm163184977
f1_keywords:
- vbawd10.chm163184977
ms.prod: word
api_name:
- Word.MailMessage.DisplayProperties
ms.assetid: fa660e11-5329-5167-ddc3-0d90ee820251
ms.date: 06/08/2017
---


# MailMessage.DisplayProperties Method (Word)

Displays the  **Properties** dialog box for the active e-mail message.


## Syntax

 _expression_ . **DisplayProperties**

 _expression_ Required. A variable that represents a **[MailMessage](mailmessage-object-word.md)** object.


## Remarks

This method is available only if you are using Word as your e-mail editor.


## Example

This example displays the  **Properties** dialog box for the active e-mail message.


```vb
Application.MailMessage.DisplayProperties
```


## See also


#### Concepts


[MailMessage Object](mailmessage-object-word.md)

