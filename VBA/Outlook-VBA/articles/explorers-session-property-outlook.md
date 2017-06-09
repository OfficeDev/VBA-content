---
title: Explorers.Session Property (Outlook)
keywords: vbaol11.chm118
f1_keywords:
- vbaol11.chm118
ms.prod: outlook
api_name:
- Outlook.Explorers.Session
ms.assetid: 51dede9c-3775-2ca9-553e-5bd87ff35ae6
ms.date: 06/08/2017
---


# Explorers.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Explorers** object.


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


[Explorers Object](explorers-object-outlook.md)

