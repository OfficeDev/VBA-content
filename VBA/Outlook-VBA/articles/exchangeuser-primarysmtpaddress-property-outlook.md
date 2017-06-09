---
title: ExchangeUser.PrimarySmtpAddress Property (Outlook)
keywords: vbaol11.chm2098
f1_keywords:
- vbaol11.chm2098
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.PrimarySmtpAddress
ms.assetid: 2dda21da-44a2-fbfe-babc-58646c76689d
ms.date: 06/08/2017
---


# ExchangeUser.PrimarySmtpAddress Property (Outlook)

Returns a  **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **PrimarySmtpAddress**

 _expression_ A variable that represents an **ExchangeUser** object.


## Remarks

This property corresponds to the MAPI property,  **PidTagEmailAddress** .

 Returns an empty string if this property has not been implemented or does not exist for the **ExchangeUser** object.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)
#### Other resources


[How to: Obtain the E-mail Address of a Recipient](http://msdn.microsoft.com/library/b645c227-a7d2-2861-3bf7-4190a19abe81%28Office.15%29.aspx)


