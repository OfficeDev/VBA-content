---
title: ExchangeDistributionList.GetMemberOfList Method (Outlook)
keywords: vbaol11.chm2130
f1_keywords:
- vbaol11.chm2130
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.GetMemberOfList
ms.assetid: daacad93-1cf4-3455-54ff-919dc4a9935e
ms.date: 06/08/2017
---


# ExchangeDistributionList.GetMemberOfList Method (Outlook)

Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains all the **[AddressEntry](addressentry-object-outlook.md)** objects representing Exchange Distribution Lists of which the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** is a member.


## Syntax

 _expression_ . **GetMemberOfList**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


### Return Value

An  **AddressEntries** collection object that represents the distribution lists of which this **ExchangeDistributionList** object is a member. Returns an **AddressEntries** object with a count of zero (0) if the **ExchangeDistributionList** is not a member of any Exchange distribution list.


## Remarks

 ** GetMemberOfList** is an expensive operation in terms of performance if there is a slow connection to Exchange Server.


## See also


#### Concepts


[ExchangeDistributionList Object](exchangedistributionlist-object-outlook.md)

