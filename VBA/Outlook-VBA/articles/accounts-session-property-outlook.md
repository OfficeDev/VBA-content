---
title: Accounts.Session Property (Outlook)
keywords: vbaol11.chm747
f1_keywords:
- vbaol11.chm747
ms.prod: outlook
api_name:
- Outlook.Accounts.Session
ms.assetid: 65be5604-6dcf-b26e-1abc-41d1a8813e90
ms.date: 06/08/2017
---


# Accounts.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Accounts** object.


## Remarks

Returns  **Null** ( **Nothing** in Visual Basic) if there is no logged-on session.

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[Accounts Object](accounts-object-outlook.md)

