---
title: Column.Session Property (Outlook)
keywords: vbaol11.chm2747
f1_keywords:
- vbaol11.chm2747
ms.prod: outlook
api_name:
- Outlook.Column.Session
ms.assetid: d0bc26d3-cb93-cc0d-ed87-9b51a2d35bcc
ms.date: 06/08/2017
---


# Column.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Column** object.


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


[Column Object](column-object-outlook.md)

