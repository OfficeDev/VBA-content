---
title: Recipients.Session Property (Outlook)
keywords: vbaol11.chm228
f1_keywords:
- vbaol11.chm228
ms.prod: outlook
api_name:
- Outlook.Recipients.Session
ms.assetid: 41ddda3c-ca79-9387-b416-8334aeecc102
ms.date: 06/08/2017
---


# Recipients.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Recipients** object.


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


[Recipients Object](recipients-object-outlook.md)

