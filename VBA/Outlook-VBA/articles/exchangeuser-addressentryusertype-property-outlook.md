---
title: ExchangeUser.AddressEntryUserType Property (Outlook)
keywords: vbaol11.chm2080
f1_keywords:
- vbaol11.chm2080
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.AddressEntryUserType
ms.assetid: fb5b16be-8846-7c9f-22bf-847d2cfc0a54
ms.date: 06/08/2017
---


# ExchangeUser.AddressEntryUserType Property (Outlook)

Returns  **olExchangeUserAddressEntry** which is a constant from the **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)** enumeration representing the user type of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **AddressEntryUserType**

 _expression_ A variable that represents an **ExchangeUser** object.


## Remarks

The  **ExchangeUser** object is derived from the **[AddressEntry](addressentry-object-outlook.md)** object. It inherits the **AddressEntryUserType** property from the **AddressEntry** object. In the case of **ExchangeUser** , **AddressEntryUserType** should always return **olExchangeUserAddressEntry** .


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

