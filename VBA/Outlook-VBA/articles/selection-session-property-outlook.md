---
title: Selection.Session Property (Outlook)
keywords: vbaol11.chm83
f1_keywords:
- vbaol11.chm83
ms.prod: outlook
api_name:
- Outlook.Selection.Session
ms.assetid: 22390a36-a51c-615d-a646-45e5aa7d253f
ms.date: 06/08/2017
---


# Selection.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Selection** object.


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


[Selection Object](selection-object-outlook.md)

