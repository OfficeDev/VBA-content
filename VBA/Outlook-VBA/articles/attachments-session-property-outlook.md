---
title: Attachments.Session Property (Outlook)
keywords: vbaol11.chm172
f1_keywords:
- vbaol11.chm172
ms.prod: outlook
api_name:
- Outlook.Attachments.Session
ms.assetid: af206370-3d50-84de-187d-019126958b61
ms.date: 06/08/2017
---


# Attachments.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Attachments** object.


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


[Attachments Object](attachments-object-outlook.md)

