---
title: ViewField.Session Property (Outlook)
keywords: vbaol11.chm2541
f1_keywords:
- vbaol11.chm2541
ms.prod: outlook
api_name:
- Outlook.ViewField.Session
ms.assetid: a6be9e3b-ebd5-410b-b0fb-f3c74b7ebd1d
ms.date: 06/08/2017
---


# ViewField.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ViewField** object.


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


[ViewField Object](viewfield-object-outlook.md)

