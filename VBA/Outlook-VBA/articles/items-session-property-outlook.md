---
title: Items.Session Property (Outlook)
keywords: vbaol11.chm55
f1_keywords:
- vbaol11.chm55
ms.prod: outlook
api_name:
- Outlook.Items.Session
ms.assetid: 5c385dfc-042e-7649-0f32-5d34e53fca57
ms.date: 06/08/2017
---


# Items.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Items** object.


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


[Items Object](items-object-outlook.md)

