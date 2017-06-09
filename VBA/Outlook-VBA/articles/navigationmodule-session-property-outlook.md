---
title: NavigationModule.Session Property (Outlook)
keywords: vbaol11.chm2805
f1_keywords:
- vbaol11.chm2805
ms.prod: outlook
api_name:
- Outlook.NavigationModule.Session
ms.assetid: 7fd04cbc-37c2-56e7-68b2-e7e8340cd99c
ms.date: 06/08/2017
---


# NavigationModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ An expression that returns a **NavigationModule** object.


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


[NavigationModule Object](navigationmodule-object-outlook.md)

