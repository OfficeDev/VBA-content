---
title: Row.Session Property (Outlook)
keywords: vbaol11.chm2241
f1_keywords:
- vbaol11.chm2241
ms.prod: outlook
api_name:
- Outlook.Row.Session
ms.assetid: a9773e62-0091-50b4-f64c-dab4217035cc
ms.date: 06/08/2017
---


# Row.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Row** object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[Row Object](row-object-outlook.md)

