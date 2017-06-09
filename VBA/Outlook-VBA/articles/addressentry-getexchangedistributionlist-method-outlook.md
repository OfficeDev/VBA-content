---
title: AddressEntry.GetExchangeDistributionList Method (Outlook)
keywords: vbaol11.chm2058
f1_keywords:
- vbaol11.chm2058
ms.prod: outlook
api_name:
- Outlook.AddressEntry.GetExchangeDistributionList
ms.assetid: 060ac302-b916-d85d-5ba8-c682894129e2
ms.date: 06/08/2017
---


# AddressEntry.GetExchangeDistributionList Method (Outlook)

Returns an  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object that represents the **[AddressEntry](addressentry-object-outlook.md)** if the **AddressEntry** belongs to an Exchange **[AddressList](addresslist-object-outlook.md)** object such as the Global Address List (GAL) and corresponds to an Exchange distribution list.


## Syntax

 _expression_ . **GetExchangeDistributionList**

 _expression_ A variable that represents an **AddressEntry** object.


### Return Value

An  **ExchangeDistributionList** object that represents the **AddressEntry** . Returns **Null** ( **Nothing** in Visual Basic) if the **AddressEntry** object does not correspond to an Exchange distribution list.


## Remarks

 You have to be connected to the Exchange server to use this method.


## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

