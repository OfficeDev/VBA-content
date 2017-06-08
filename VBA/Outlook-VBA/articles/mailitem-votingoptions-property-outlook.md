---
title: MailItem.VotingOptions Property (Outlook)
keywords: vbaol11.chm1363
f1_keywords:
- vbaol11.chm1363
ms.prod: outlook
api_name:
- Outlook.MailItem.VotingOptions
ms.assetid: 696b6dfe-1840-d43b-e6ec-e410a387665c
ms.date: 06/08/2017
---


# MailItem.VotingOptions Property (Outlook)

Returns or sets a  **String** specifying a delimited string containing the voting options for the mail message. Read/write.


## Syntax

 _expression_ . **VotingOptions**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property uses the character specified in the value name,  **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple voting options.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

