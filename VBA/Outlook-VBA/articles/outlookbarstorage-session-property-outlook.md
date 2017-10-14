---
title: OutlookBarStorage.Session Property (Outlook)
keywords: vbaol11.chm370
f1_keywords:
- vbaol11.chm370
ms.prod: outlook
api_name:
- Outlook.OutlookBarStorage.Session
ms.assetid: f3ba6302-aca2-f8ba-3a82-ae35f6b5b609
ms.date: 06/08/2017
---


# OutlookBarStorage.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **OutlookBarStorage** object.


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


[OutlookBarStorage Object](outlookbarstorage-object-outlook.md)

