---
title: Account.Session Property (Outlook)
keywords: vbaol11.chm738
f1_keywords:
- vbaol11.chm738
ms.prod: outlook
api_name:
- Outlook.Account.Session
ms.assetid: 92890235-402c-80c8-10b7-7339f153134e
ms.date: 06/08/2017
---


# Account.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Account** object.


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


[Account Object](account-object-outlook.md)

