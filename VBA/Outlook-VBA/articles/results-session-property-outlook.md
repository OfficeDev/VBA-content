---
title: Results.Session Property (Outlook)
keywords: vbaol11.chm500
f1_keywords:
- vbaol11.chm500
ms.prod: outlook
api_name:
- Outlook.Results.Session
ms.assetid: 3b6453fb-ba9e-b0c1-f559-f751cbe142e2
ms.date: 06/08/2017
---


# Results.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Results** object.


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


[Results Object](results-object-outlook.md)

