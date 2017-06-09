---
title: Account.SmtpAddress Property (Outlook)
keywords: vbaol11.chm743
f1_keywords:
- vbaol11.chm743
ms.prod: outlook
api_name:
- Outlook.Account.SmtpAddress
ms.assetid: 443beb7a-0ada-8e86-69d7-63880033abca
ms.date: 06/08/2017
---


# Account.SmtpAddress Property (Outlook)

Returns a  **String** representing the Simple Mail Transfer Protocol (SMTP) address for the **[Account](account-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **SmtpAddress**

 _expression_ A variable that represents an **Account** object.


## Remarks

The purpose of  **SmtpAddress** and **[Account.UserName](account-username-property-outlook.md)** is to provide an account-based context to determine identity.

If the account does not have an SMTP address,  **SmtpAddress** returns an empty string.


## See also


#### Concepts


[Account Object](account-object-outlook.md)
#### Other resources


[Send an E-mail Given the SMTP Address of an Account](send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)



