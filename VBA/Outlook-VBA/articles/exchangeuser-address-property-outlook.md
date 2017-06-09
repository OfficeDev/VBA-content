---
title: ExchangeUser.Address Property (Outlook)
keywords: vbaol11.chm2065
f1_keywords:
- vbaol11.chm2065
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Address
ms.assetid: b3a36b16-e652-9e3f-86fd-7cea0c72d78c
ms.date: 06/08/2017
---


# ExchangeUser.Address Property (Outlook)

Returns or sets a  **String** representing the X400 e-mail address of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **Address**

 _expression_ A variable that represents an **ExchangeUser** object.


## Remarks

This property assumes the X400 address of the user. To determine the primary Internet address, use the  **[ExchangeUser.PrimarySmtpAddress](exchangeuser-primarysmtpaddress-property-outlook.md)** property.

The  **Address** property must be set before calling the **[ExchangeUser.Details](exchangeuser-details-method-outlook.md)** method.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

