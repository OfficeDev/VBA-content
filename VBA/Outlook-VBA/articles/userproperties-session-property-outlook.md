---
title: UserProperties.Session Property (Outlook)
keywords: vbaol11.chm205
f1_keywords:
- vbaol11.chm205
ms.prod: outlook
api_name:
- Outlook.UserProperties.Session
ms.assetid: 0cd76318-80c6-4cfc-3aca-32e385ff6b88
ms.date: 06/08/2017
---


# UserProperties.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **UserProperties** object.


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


[UserProperties Object](userproperties-object-outlook.md)

