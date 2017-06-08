---
title: Search.Session Property (Outlook)
keywords: vbaol11.chm2251
f1_keywords:
- vbaol11.chm2251
ms.prod: outlook
api_name:
- Outlook.Search.Session
ms.assetid: 8d5a2300-dc21-0fbe-c7c0-17741caae30a
ms.date: 06/08/2017
---


# Search.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Search** object.


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


[Search Object](search-object-outlook.md)

