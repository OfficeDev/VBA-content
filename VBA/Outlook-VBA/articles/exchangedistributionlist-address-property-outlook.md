---
title: ExchangeDistributionList.Address Property (Outlook)
keywords: vbaol11.chm2112
f1_keywords:
- vbaol11.chm2112
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Address
ms.assetid: 9bfb7b5c-02ec-febc-c411-574efaa52c55
ms.date: 06/08/2017
---


# ExchangeDistributionList.Address Property (Outlook)

Returns or sets a  **String** representing the X400 e-mail address of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **Address**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


## Remarks

This property assumes the X400 address of the distribution list. To determine the primary Internet address, use the  **[ExchangeDistributionList.PrimarySmtpAddress](exchangedistributionlist-primarysmtpaddress-property-outlook.md)** property.

The  **Address** property must be set before calling the **[ExchangeDistributionList.Details](exchangeuser-details-method-outlook.md)** method.


## See also


#### Concepts


[ExchangeDistributionList Object](exchangedistributionlist-object-outlook.md)

