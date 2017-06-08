---
title: Exception.Session Property (Outlook)
keywords: vbaol11.chm299
f1_keywords:
- vbaol11.chm299
ms.prod: outlook
api_name:
- Outlook.Exception.Session
ms.assetid: b8663ef0-1042-e3c4-81ca-76d4b76a3351
ms.date: 06/08/2017
---


# Exception.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Exception** object.


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


[Exception Object](exception-object-outlook.md)

