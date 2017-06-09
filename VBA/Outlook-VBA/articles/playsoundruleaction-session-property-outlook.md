---
title: PlaySoundRuleAction.Session Property (Outlook)
keywords: vbaol11.chm2273
f1_keywords:
- vbaol11.chm2273
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction.Session
ms.assetid: 8d3e9f6e-848d-9879-61a8-7662858674d4
ms.date: 06/08/2017
---


# PlaySoundRuleAction.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **PlaySoundRuleAction** object.


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


[PlaySoundRuleAction Object](playsoundruleaction-object-outlook.md)

