---
title: FormRegion.Session Property (Outlook)
keywords: vbaol11.chm2387
f1_keywords:
- vbaol11.chm2387
ms.prod: outlook
api_name:
- Outlook.FormRegion.Session
ms.assetid: 13b9a148-c898-a3ef-8341-073767ce665e
ms.date: 06/08/2017
---


# FormRegion.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **FormRegion** object.


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


[FormRegion Object](formregion-object-outlook.md)

