---
title: AddressEntry.Session Property (Outlook)
keywords: vbaol11.chm2040
f1_keywords:
- vbaol11.chm2040
ms.prod: outlook
api_name:
- Outlook.AddressEntry.Session
ms.assetid: e2fdc0ed-a470-eca7-0709-ea7938df3516
ms.date: 06/08/2017
---


# AddressEntry.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AddressEntry** object.


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


[AddressEntry Object](addressentry-object-outlook.md)

