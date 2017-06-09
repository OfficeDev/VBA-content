---
title: NavigationModules.Session Property (Outlook)
keywords: vbaol11.chm2797
f1_keywords:
- vbaol11.chm2797
ms.prod: outlook
api_name:
- Outlook.NavigationModules.Session
ms.assetid: ce7f293c-cce6-5471-fd41-3387c2f0195e
ms.date: 06/08/2017
---


# NavigationModules.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ An expression that returns a **NavigationModules** object.


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


[NavigationModules Object](navigationmodules-object-outlook.md)

