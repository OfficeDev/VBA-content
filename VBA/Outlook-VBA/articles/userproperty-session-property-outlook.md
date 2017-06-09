---
title: UserProperty.Session Property (Outlook)
keywords: vbaol11.chm215
f1_keywords:
- vbaol11.chm215
ms.prod: outlook
api_name:
- Outlook.UserProperty.Session
ms.assetid: 181d0aad-9b03-9cce-b6dd-33a290d57ee9
ms.date: 06/08/2017
---


# UserProperty.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **UserProperty** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[UserProperty Object](userproperty-object-outlook.md)

