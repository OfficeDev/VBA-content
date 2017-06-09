---
title: AddressEntries.Session Property (Outlook)
keywords: vbaol11.chm27
f1_keywords:
- vbaol11.chm27
ms.prod: outlook
api_name:
- Outlook.AddressEntries.Session
ms.assetid: bdd2afb2-a4f7-e31b-9413-94ba7e6ca213
ms.date: 06/08/2017
---


# AddressEntries.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AddressEntries** object.


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


[AddressEntries Object](addressentries-object-outlook.md)

