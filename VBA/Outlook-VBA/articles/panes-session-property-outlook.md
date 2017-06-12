---
title: Panes.Session Property (Outlook)
keywords: vbaol11.chm76
f1_keywords:
- vbaol11.chm76
ms.prod: outlook
api_name:
- Outlook.Panes.Session
ms.assetid: 3f0eeae2-e02e-d7f1-70de-6c9d869756d9
ms.date: 06/08/2017
---


# Panes.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Panes** object.


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


[Panes Object](panes-object-outlook.md)

