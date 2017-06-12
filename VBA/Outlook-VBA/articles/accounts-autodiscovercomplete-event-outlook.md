---
title: Accounts.AutoDiscoverComplete Event (Outlook)
keywords: vbaol11.chm3477
f1_keywords:
- vbaol11.chm3477
ms.prod: outlook
api_name:
- Outlook.Accounts.AutoDiscoverComplete
ms.assetid: 86738163-4fb3-b2f5-40bd-4704081d4564
ms.date: 06/08/2017
---


# Accounts.AutoDiscoverComplete Event (Outlook)

Occurs after Microsoft Outlook has finished accessing the auto-discovery service of the Microsoft Exchange Server that is associated with the account, and has the related information available in the  **[AutoDiscoverXml](account-autodiscoverxml-property-outlook.md)** property of the **[Account](account-object-outlook.md)** object.


## Syntax

 _expression_ . **AutoDiscoverComplete**( **_Account_** )

 _expression_ A variable that represents an **[Accounts](accounts-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required| **Account**|The account whose auto-discovery of the associated Exchange Server is complete.|

## Remarks

This event is similar to the  **[AutoDiscoverComplete](namespace-autodiscovercomplete-event-outlook.md)** event of the **[NameSpace](namespace-object-outlook.md)** object, except that this event applies to the account for which auto-discovery is completed and not necessarily to the primary Exchange account.


## See also


#### Concepts


[Accounts Object](accounts-object-outlook.md)

