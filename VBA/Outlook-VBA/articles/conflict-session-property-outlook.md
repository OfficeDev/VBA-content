---
title: Conflict.Session Property (Outlook)
keywords: vbaol11.chm413
f1_keywords:
- vbaol11.chm413
ms.prod: outlook
api_name:
- Outlook.Conflict.Session
ms.assetid: cd7eaf1e-545b-5a40-d95c-841f72a7a15e
ms.date: 06/08/2017
---


# Conflict.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Conflict** object.


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


[Conflict Object](conflict-object-outlook.md)

