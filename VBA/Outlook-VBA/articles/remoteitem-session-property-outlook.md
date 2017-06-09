---
title: RemoteItem.Session Property (Outlook)
keywords: vbaol11.chm1584
f1_keywords:
- vbaol11.chm1584
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Session
ms.assetid: 2692f2ef-b8cb-1b0e-25fb-0381f98c7e79
ms.date: 06/08/2017
---


# RemoteItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **RemoteItem** object.


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


[RemoteItem Object](remoteitem-object-outlook.md)

