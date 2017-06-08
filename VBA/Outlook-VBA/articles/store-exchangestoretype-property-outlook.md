---
title: Store.ExchangeStoreType Property (Outlook)
keywords: vbaol11.chm802
f1_keywords:
- vbaol11.chm802
ms.prod: outlook
api_name:
- Outlook.Store.ExchangeStoreType
ms.assetid: ca6002bd-444d-a111-adca-6f8fafc37ea1
ms.date: 06/08/2017
---


# Store.ExchangeStoreType Property (Outlook)

Returns a constant in the  **[OlExchangeStoreType](olexchangestoretype-enumeration-outlook.md)** enumeration that indicates the type of an Exchange store. Read-only.


## Syntax

 _expression_ . **ExchangeStoreType**

 _expression_ A variable that represents a **Store** object.


## Remarks

The  **ExchangeStoreType** property distinguishes among different Exchange store types, such as primary Exchange mailbox, Exchange mailbox, Public Folder store, or non-Exchange store. This property does not distinguish among every type of store including Hotmail, HTTP, IMAP, and so forth. Use **[Account.AccountType](account-accounttype-property-outlook.md)** for the type of server associated with an e-mail account, such as Exchange, HTTP, IMAP, or POP3.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

