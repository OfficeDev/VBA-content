---
title: Recipient.Session Property (Outlook)
keywords: vbaol11.chm2342
f1_keywords:
- vbaol11.chm2342
ms.prod: outlook
api_name:
- Outlook.Recipient.Session
ms.assetid: 0719e438-c9b0-ecca-1aa0-f25c9b21fe69
ms.date: 06/08/2017
---


# Recipient.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Recipient** object.


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


[Recipient Object](recipient-object-outlook.md)

