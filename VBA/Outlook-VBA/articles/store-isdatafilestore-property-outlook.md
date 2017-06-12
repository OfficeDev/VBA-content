---
title: Store.IsDataFileStore Property (Outlook)
keywords: vbaol11.chm805
f1_keywords:
- vbaol11.chm805
ms.prod: outlook
api_name:
- Outlook.Store.IsDataFileStore
ms.assetid: 76dc73b7-1d19-465f-744f-1209211f2496
ms.date: 06/08/2017
---


# Store.IsDataFileStore Property (Outlook)

Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is a store for an Outlook data file, which is either a Personal Folders File (.pst) or an Offline Folder File (.ost). Read-only.


## Syntax

 _expression_ . **IsDataFileStore**

 _expression_ A variable that represents a **Store** object.


## Remarks

 **IsDataFileStore** supports only Exchange stores, and will return **False** for HTTP-type stores such as Hotmail and MSN, and for IMAP stores.

For Exchange stores,  **IsDataFileStore** will return **False** if the user profile is not using Cached Exchange mode. **IsDataFileStore** will also return **False** when the store is an Exchange Public Folder (that is, **[Store.ExchangeStoreType](store-exchangestoretype-property-outlook.md)** is **olExchangePublicFolder** ).

 **IsDataFileStore** does not indicate whether the store is located on a local hard drive. For example, a .pst file could be located on a mapped network drive and **IsDataFileStore** would still return **True** .

The return value of  **IsDataFileStore** can change if the user is configured for classic Exchange offline mode. When the user is offline and using classic Exchange offline mode, **IsDataFileStore** will return **True** . When the user is online and using classic Exchange online mode, **IsDataFileStore** will return **False** .


## See also


#### Concepts


[Store Object](store-object-outlook.md)

