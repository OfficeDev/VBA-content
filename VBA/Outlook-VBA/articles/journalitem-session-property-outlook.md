---
title: JournalItem.Session Property (Outlook)
keywords: vbaol11.chm1229
f1_keywords:
- vbaol11.chm1229
ms.prod: outlook
api_name:
- Outlook.JournalItem.Session
ms.assetid: d691078d-f651-c31a-d767-0b3bd91df800
ms.date: 06/08/2017
---


# JournalItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **JournalItem** object.


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


[JournalItem Object](journalitem-object-outlook.md)

