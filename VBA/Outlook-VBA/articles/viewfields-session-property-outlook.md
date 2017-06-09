---
title: ViewFields.Session Property (Outlook)
keywords: vbaol11.chm2548
f1_keywords:
- vbaol11.chm2548
ms.prod: outlook
api_name:
- Outlook.ViewFields.Session
ms.assetid: 480ac826-b966-9204-8850-214b53a1b0da
ms.date: 06/08/2017
---


# ViewFields.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ViewFields** object.


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


[ViewFields Object](viewfields-object-outlook.md)

