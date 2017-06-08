---
title: OutlookBarGroups.Session Property (Outlook)
keywords: vbaol11.chm348
f1_keywords:
- vbaol11.chm348
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups.Session
ms.assetid: f62d8290-7e42-1dbb-1135-3298b47124d6
ms.date: 06/08/2017
---


# OutlookBarGroups.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **OutlookBarGroups** object.


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


[OutlookBarGroups Object](outlookbargroups-object-outlook.md)

