---
title: NameSpace.GetGlobalAddressList Method (Outlook)
keywords: vbaol11.chm785
f1_keywords:
- vbaol11.chm785
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetGlobalAddressList
ms.assetid: 0c892483-96c5-461d-a862-fe84ddcce097
ms.date: 06/08/2017
---


# NameSpace.GetGlobalAddressList Method (Outlook)

Returns an  **[AddressList](addresslist-object-outlook.md)** object that represents the Exchange Global Address List.


## Syntax

 _expression_ . **GetGlobalAddressList**

 _expression_ A variable that represents a **NameSpace** object.


### Return Value

An  **AddressList** that represents the Global Address List.


## Remarks

 **GetGlobalAddressList** supports only Exchange servers. It returns an error if the Global Address List is not available or cannot be found.

It also returns an error if no connection is available or the user is set to work offline.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

