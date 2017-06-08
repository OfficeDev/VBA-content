---
title: ContactsModule.Session Property (Outlook)
keywords: vbaol11.chm2834
f1_keywords:
- vbaol11.chm2834
ms.prod: outlook
api_name:
- Outlook.ContactsModule.Session
ms.assetid: 4ab5d6e1-fcff-9aa4-0779-a365e01d0a1c
ms.date: 06/08/2017
---


# ContactsModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ContactsModule** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[ContactsModule Object](contactsmodule-object-outlook.md)

