---
title: OutlookBarPane.Session Property (Outlook)
keywords: vbaol11.chm361
f1_keywords:
- vbaol11.chm361
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.Session
ms.assetid: 8aa3d36b-2044-85a7-2b79-86c6b161c4df
ms.date: 06/08/2017
---


# OutlookBarPane.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **OutlookBarPane** object.


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


[OutlookBarPane Object](outlookbarpane-object-outlook.md)

