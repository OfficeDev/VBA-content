---
title: ExchangeDistributionList.PrimarySmtpAddress Property (Outlook)
keywords: vbaol11.chm2134
f1_keywords:
- vbaol11.chm2134
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.PrimarySmtpAddress
ms.assetid: f64bbc29-14c4-be68-402a-16d9ac34a727
ms.date: 06/08/2017
---


# ExchangeDistributionList.PrimarySmtpAddress Property (Outlook)

Returns a  **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **PrimarySmtpAddress**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


## Remarks

This property corresponds to the MAPI property  **PidTagEmailAddress** .

Returns an empty string if this property has not been implemented or does not exist for the  **ExchangeDistributionList** object.


## See also


#### Concepts


[ExchangeDistributionList Object](exchangedistributionlist-object-outlook.md)
#### Other resources


[How to: Obtain the E-mail Address of a Recipient](http://msdn.microsoft.com/library/b645c227-a7d2-2861-3bf7-4190a19abe81%28Office.15%29.aspx)


