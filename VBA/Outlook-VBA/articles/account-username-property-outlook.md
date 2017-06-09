---
title: Account.UserName Property (Outlook)
keywords: vbaol11.chm742
f1_keywords:
- vbaol11.chm742
ms.prod: outlook
api_name:
- Outlook.Account.UserName
ms.assetid: 3ab96240-b68c-e2f7-83b9-6d6663c4880d
ms.date: 06/08/2017
---


# Account.UserName Property (Outlook)

Returns a  **String** representing the user name for the **[Account](account-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **UserName**

 _expression_ A variable that represents an **Account** object.


## Remarks

The purpose of  **[Account.SmtpAddress](account-smtpaddress-property-outlook.md)** and **UserName** is to provide an account-based context to determine identity.

If the account does not have a user name defined,  **UserName** returns an empty string.


## See also


#### Concepts


[Account Object](account-object-outlook.md)

