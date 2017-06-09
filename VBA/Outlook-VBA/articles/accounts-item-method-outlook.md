---
title: Accounts.Item Method (Outlook)
keywords: vbaol11.chm750
f1_keywords:
- vbaol11.chm750
ms.prod: outlook
api_name:
- Outlook.Accounts.Item
ms.assetid: 8ef9c358-6d8b-1cbb-40ed-6d3462ae335e
ms.date: 06/08/2017
---


# Accounts.Item Method (Outlook)

Returns an  **[Account](account-object-outlook.md)** object specified by _Index_ .


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **Accounts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A one-based  **Long** that indexes into the **[Accounts](accounts-object-outlook.md)** collection, or a **String** that specifies the **[DisplayName](account-displayname-property-outlook.md)** of an **Account** .|

### Return Value

An  **Account** object that matches the account specified by _Index_ .


## See also


#### Concepts


[Accounts Object](accounts-object-outlook.md)

