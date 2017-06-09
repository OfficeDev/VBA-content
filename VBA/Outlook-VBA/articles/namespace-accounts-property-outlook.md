---
title: NameSpace.Accounts Property (Outlook)
keywords: vbaol11.chm778
f1_keywords:
- vbaol11.chm778
ms.prod: outlook
api_name:
- Outlook.NameSpace.Accounts
ms.assetid: 80e969ea-d2cc-966d-5fe4-68d59951b5c9
ms.date: 06/08/2017
---


# NameSpace.Accounts Property (Outlook)

Returns an  **[Accounts](accounts-object-outlook.md)** collection object that represents all the **[Account](account-object-outlook.md)** objects in the current profile. Read-only.


## Syntax

 _expression_ . **Accounts**

 _expression_ A variable that represents a **NameSpace** object.


## Remarks

If Outlook is running in sessionless mode,  **Accounts** returns an **Accounts** collection with **[Accounts.Count](accounts-count-property-outlook.md)** equal to 0.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

