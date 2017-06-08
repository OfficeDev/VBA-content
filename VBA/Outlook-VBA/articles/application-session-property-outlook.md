---
title: Application.Session Property (Outlook)
keywords: vbaol11.chm707
f1_keywords:
- vbaol11.chm707
ms.prod: outlook
api_name:
- Outlook.Application.Session
ms.assetid: 720b2849-fe01-afb3-363c-f3bf0cd7d872
ms.date: 06/08/2017
---


# Application.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Application** object.


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


[Application Object](application-object-outlook.md)

