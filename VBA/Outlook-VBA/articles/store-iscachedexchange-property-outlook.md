---
title: Store.IsCachedExchange Property (Outlook)
keywords: vbaol11.chm804
f1_keywords:
- vbaol11.chm804
ms.prod: outlook
api_name:
- Outlook.Store.IsCachedExchange
ms.assetid: 2f3fbd5d-8cf1-5fdd-6074-f4da4216dcd4
ms.date: 06/08/2017
---


# Store.IsCachedExchange Property (Outlook)

Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is a cached Exchange store. Read-only.


## Syntax

 _expression_ . **IsCachedExchange**

 _expression_ A variable that represents a **Store** object.


## Remarks

 **IsCachedExchange** returns **True** when **[Store.ExchangeStoreType](store-exchangestoretype-property-outlook.md)** is a primary Exchange mailbox ( **Store.ExchangeStoreType** is **olExchangePrimaryMailbox** ) and the mailbox is configured to use cached Exchange mode. It returns **False** otherwise. In particular, it returns **False** for an Exchange Public Folder store that is configured to cache Public Folder favorites.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

