---
title: Action.Session Property (Outlook)
keywords: vbaol11.chm12
f1_keywords:
- vbaol11.chm12
ms.prod: outlook
api_name:
- Outlook.Action.Session
ms.assetid: cfe619d2-3a7e-c8af-de17-be2363de0a56
ms.date: 06/08/2017
---


# Action.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Action** object.


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


[Action Object](action-object-outlook.md)

