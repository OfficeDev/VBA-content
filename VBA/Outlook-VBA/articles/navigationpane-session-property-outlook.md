---
title: NavigationPane.Session Property (Outlook)
keywords: vbaol11.chm2788
f1_keywords:
- vbaol11.chm2788
ms.prod: outlook
api_name:
- Outlook.NavigationPane.Session
ms.assetid: 038fd9d2-77e3-3af2-b8f5-b491b6e4f2ab
ms.date: 06/08/2017
---


# NavigationPane.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **NavigationPane** object.


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


[NavigationPane Object](navigationpane-object-outlook.md)

