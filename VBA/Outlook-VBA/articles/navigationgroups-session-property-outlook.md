---
title: NavigationGroups.Session Property (Outlook)
keywords: vbaol11.chm2854
f1_keywords:
- vbaol11.chm2854
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Session
ms.assetid: b742bee6-7067-8168-ebd9-2823da65dd0f
ms.date: 06/08/2017
---


# NavigationGroups.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **NavigationGroups** object.


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


[NavigationGroups Object](navigationgroups-object-outlook.md)

