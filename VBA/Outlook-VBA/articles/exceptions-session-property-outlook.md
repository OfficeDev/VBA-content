---
title: Exceptions.Session Property (Outlook)
keywords: vbaol11.chm292
f1_keywords:
- vbaol11.chm292
ms.prod: outlook
api_name:
- Outlook.Exceptions.Session
ms.assetid: a0674664-e087-3df7-b80a-419863255160
ms.date: 06/08/2017
---


# Exceptions.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Exceptions** object.


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


[Exceptions Object](exceptions-object-outlook.md)

