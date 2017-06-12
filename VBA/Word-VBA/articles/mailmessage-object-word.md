---
title: MailMessage Object (Word)
keywords: vbawd10.chm2490
f1_keywords:
- vbawd10.chm2490
ms.prod: word
api_name:
- Word.MailMessage
ms.assetid: d0109969-27f7-0180-c56d-5b49a3f0171b
ms.date: 06/08/2017
---


# MailMessage Object (Word)

Represents the active e-mail message if you are using Microsoft Word as your e-mail editor.


## Remarks

Use the  **MailMessage** property to return the **MailMessage** object. The following example validates the e-mail addresses that appear in the active e-mail message.


```vb
Application.MailMessage.CheckName
```

The methods of the  **MailMessage** object require that you are using Word as your e-mail editor and that an e-mail message is active. If either of these conditions is not true, an error occurs.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

