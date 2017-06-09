---
title: View.Session Property (Outlook)
keywords: vbaol11.chm2482
f1_keywords:
- vbaol11.chm2482
ms.prod: outlook
api_name:
- Outlook.View.Session
ms.assetid: 32c6c27e-2351-c10c-47cd-bcca06d25660
ms.date: 06/08/2017
---


# View.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **View** object.


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


[View Object](view-object-outlook.md)

