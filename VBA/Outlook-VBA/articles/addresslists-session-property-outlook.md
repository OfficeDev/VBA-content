---
title: AddressLists.Session Property (Outlook)
keywords: vbaol11.chm90
f1_keywords:
- vbaol11.chm90
ms.prod: outlook
api_name:
- Outlook.AddressLists.Session
ms.assetid: 60b4307f-92c7-abed-5bc7-2a190cddd4ca
ms.date: 06/08/2017
---


# AddressLists.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AddressLists** object.


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


[AddressLists Object](addresslists-object-outlook.md)

