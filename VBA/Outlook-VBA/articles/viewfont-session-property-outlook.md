---
title: ViewFont.Session Property (Outlook)
keywords: vbaol11.chm2693
f1_keywords:
- vbaol11.chm2693
ms.prod: outlook
api_name:
- Outlook.ViewFont.Session
ms.assetid: 8f126189-3bec-6eee-1e62-b178738d361b
ms.date: 06/08/2017
---


# ViewFont.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ViewFont** object.


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


[ViewFont Object](viewfont-object-outlook.md)

