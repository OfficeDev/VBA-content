---
title: SyncObjects.Session Property (Outlook)
keywords: vbaol11.chm97
f1_keywords:
- vbaol11.chm97
ms.prod: outlook
api_name:
- Outlook.SyncObjects.Session
ms.assetid: 443c2e6d-fda7-8230-b3b1-bd87cccafe23
ms.date: 06/08/2017
---


# SyncObjects.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **SyncObjects** object.


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


[SyncObjects Object](syncobjects-object-outlook.md)

