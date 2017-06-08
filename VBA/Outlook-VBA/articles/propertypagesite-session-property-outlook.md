---
title: PropertyPageSite.Session Property (Outlook)
keywords: vbaol11.chm387
f1_keywords:
- vbaol11.chm387
ms.prod: outlook
api_name:
- Outlook.PropertyPageSite.Session
ms.assetid: 0e1dd77d-fcd8-afe7-7370-3b755c910452
ms.date: 06/08/2017
---


# PropertyPageSite.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **PropertyPageSite** object.


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


[PropertyPageSite Object](propertypagesite-object-outlook.md)

