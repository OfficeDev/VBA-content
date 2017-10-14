---
title: MoveOrCopyRuleAction.Session Property (Outlook)
keywords: vbaol11.chm2210
f1_keywords:
- vbaol11.chm2210
ms.prod: outlook
api_name:
- Outlook.MoveOrCopyRuleAction.Session
ms.assetid: a77c9ccf-6f8d-92de-f6d4-1f3d7e92c810
ms.date: 06/08/2017
---


# MoveOrCopyRuleAction.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **MoveOrCopyRuleAction** object.


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


[MoveOrCopyRuleAction Object](moveorcopyruleaction-object-outlook.md)

