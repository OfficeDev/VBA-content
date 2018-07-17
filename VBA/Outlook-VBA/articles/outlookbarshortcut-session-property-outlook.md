---
title: OutlookBarShortcut.Session Property (Outlook)
keywords: vbaol11.chm340
f1_keywords:
- vbaol11.chm340
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut.Session
ms.assetid: aee32453-1650-1d28-10ae-a125fa4c4394
ms.date: 06/08/2017
---


# OutlookBarShortcut.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **OutlookBarShortcut** object.


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


[OutlookBarShortcut Object](outlookbarshortcut-object-outlook.md)

