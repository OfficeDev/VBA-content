---
title: Stores.Session Property (Outlook)
keywords: vbaol11.chm816
f1_keywords:
- vbaol11.chm816
ms.prod: outlook
api_name:
- Outlook.Stores.Session
ms.assetid: aea9466c-4b22-10fa-7938-d12f4f193148
ms.date: 06/08/2017
---


# Stores.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Stores** object.


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


[Stores Object](stores-object-outlook.md)

