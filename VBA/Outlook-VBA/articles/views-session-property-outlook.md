---
title: Views.Session Property (Outlook)
keywords: vbaol11.chm543
f1_keywords:
- vbaol11.chm543
ms.prod: outlook
api_name:
- Outlook.Views.Session
ms.assetid: 677d7b97-b138-3506-7b45-26d091f9ba6e
ms.date: 06/08/2017
---


# Views.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Views** object.


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


[Views Object](views-object-outlook.md)

