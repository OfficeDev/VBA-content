---
title: AddressList.Session Property (Outlook)
keywords: vbaol11.chm2025
f1_keywords:
- vbaol11.chm2025
ms.prod: outlook
api_name:
- Outlook.AddressList.Session
ms.assetid: ac7d208a-49c8-fe1a-ea33-f7c6d8a700d7
ms.date: 06/08/2017
---


# AddressList.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AddressList** object.


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


[AddressList Object](addresslist-object-outlook.md)

