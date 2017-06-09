---
title: ExchangeUser.GetDirectReports Method (Outlook)
keywords: vbaol11.chm2083
f1_keywords:
- vbaol11.chm2083
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetDirectReports
ms.assetid: 753201ad-8001-3185-7d68-fda15907099d
ms.date: 06/08/2017
---


# ExchangeUser.GetDirectReports Method (Outlook)

Obtains an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains all the users directly reporting to the Exchange user.


## Syntax

 _expression_ . **GetDirectReports**

 _expression_ A variable that represents an **ExchangeUser** object.


### Return Value

An  **AddressEntries** collection object that contains the users directly reporting to the Exchange user. The **AddressEntries** object will have a count of zero (0) if there is no direct report represented by an **[AddressEntry](addressentry-object-outlook.md)** in the current session, or if direct reports have not been implemented in the Exchange directory.


## Remarks

 **GetDirectReports** is an expensive operation in terms of performance if there is a slow connection to the Exchange server.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

