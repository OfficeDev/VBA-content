---
title: DistListItem.Session Property (Outlook)
keywords: vbaol11.chm1112
f1_keywords:
- vbaol11.chm1112
ms.prod: outlook
api_name:
- Outlook.DistListItem.Session
ms.assetid: c36e7ef0-baf0-daa3-2e9a-c8d9d78bb6d5
ms.date: 06/08/2017
---


# DistListItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **DistListItem** object.


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


[DistListItem Object](distlistitem-object-outlook.md)

