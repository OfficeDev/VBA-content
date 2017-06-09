---
title: AccountSelector.Session Property (Outlook)
keywords: vbaol11.chm3451
f1_keywords:
- vbaol11.chm3451
ms.prod: outlook
api_name:
- Outlook.AccountSelector.Session
ms.assetid: aca5ce47-5597-8bb3-588b-0c58d704b158
ms.date: 06/08/2017
---


# AccountSelector.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **[AccountSelector](accountselector-object-outlook.md)** object.


## Remarks

Returns  **Null** ( **Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method of the **[Application](application-object-outlook.md)** object interchangeably to obtain the **NameSpace** object for the current session.


## See also


#### Concepts


[AccountSelector Object](accountselector-object-outlook.md)

