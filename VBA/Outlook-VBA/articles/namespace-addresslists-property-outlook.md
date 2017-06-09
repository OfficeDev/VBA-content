---
title: NameSpace.AddressLists Property (Outlook)
keywords: vbaol11.chm759
f1_keywords:
- vbaol11.chm759
ms.prod: outlook
api_name:
- Outlook.NameSpace.AddressLists
ms.assetid: 68b236db-f964-6f7f-6246-e79c6ada19e9
ms.date: 06/08/2017
---


# NameSpace.AddressLists Property (Outlook)

Returns an  **[AddressLists](addresslists-object-outlook.md)** collection representing a collection of the address lists available for this session. Read-only.


## Syntax

 _expression_ . **AddressLists**

 _expression_ A variable that represents a **NameSpace** object.


## Remarks

The  **AddressLists** collection represents the root of the address book hierarchy for the current session. A particular **[AddressList](addresslist-object-outlook.md)** object represents one of the available address books. The type of access you obtain depends on the access permissions granted to you by each individual address book provider.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

