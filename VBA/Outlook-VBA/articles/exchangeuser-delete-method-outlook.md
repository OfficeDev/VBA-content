---
title: ExchangeUser.Delete Method (Outlook)
keywords: vbaol11.chm2073
f1_keywords:
- vbaol11.chm2073
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Delete
ms.assetid: d11a82c4-28de-efef-5170-20f999f2bf08
ms.date: 06/08/2017
---


# ExchangeUser.Delete Method (Outlook)

Deletes the  **[ExchangeUser](exchangeuser-object-outlook.md)** object from the **[AddressEntries](addressentries-object-outlook.md)** collection object to which it belongs.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents an **ExchangeUser** object.


## Remarks

The  **ExchangeUser** object is derived from the **[AddressEntry](addressentry-object-outlook.md)** object. An **ExchangeUser** object is an **AddressEntry** object that has **olExchangeUserAddressEntry** as the **[AddressEntry.AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)** ; calling **[AddressEntry.GetExchangeUser](addressentry-getexchangeuser-method-outlook.md)** returns the corresponding **ExchangeUser** object.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

