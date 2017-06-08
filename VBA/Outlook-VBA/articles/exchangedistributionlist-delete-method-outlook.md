---
title: ExchangeDistributionList.Delete Method (Outlook)
keywords: vbaol11.chm2120
f1_keywords:
- vbaol11.chm2120
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Delete
ms.assetid: f1d14d2f-63ba-d02a-d40f-56f7d807e11e
ms.date: 06/08/2017
---


# ExchangeDistributionList.Delete Method (Outlook)

Deletes the  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object from the **[AddressEntries](addressentries-object-outlook.md)** collection object to which it belongs.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


## Remarks

The  **ExchangeDistributionList** object is derived from the **[AddressEntry](addressentry-object-outlook.md)** object. An **ExchangeDistributionList** object is an **AddressEntry** object that has **olExchangeDistributionListAddressEntry** as the **[AddressEntry.AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)** ; calling **[AddressEntry.GetExchangeDistributionList](addressentry-getexchangedistributionlist-method-outlook.md)** returns the corresponding **ExchangeDistributionList** object.


## See also


#### Concepts


[ExchangeDistributionList Object](exchangedistributionlist-object-outlook.md)

