---
title: ColumnFormat.Session Property (Outlook)
keywords: vbaol11.chm2726
f1_keywords:
- vbaol11.chm2726
ms.prod: outlook
api_name:
- Outlook.ColumnFormat.Session
ms.assetid: 6836c80e-5194-0a90-477f-3ed51a91c3b6
ms.date: 06/08/2017
---


# ColumnFormat.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ColumnFormat** object.


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


[ColumnFormat Object](columnformat-object-outlook.md)

