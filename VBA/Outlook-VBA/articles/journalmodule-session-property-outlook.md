---
title: JournalModule.Session Property (Outlook)
keywords: vbaol11.chm2864
f1_keywords:
- vbaol11.chm2864
ms.prod: outlook
api_name:
- Outlook.JournalModule.Session
ms.assetid: 416b232d-bed3-fcf5-db47-2946b5a8d244
ms.date: 06/08/2017
---


# JournalModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **JournalModule** object.


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


[JournalModule Object](journalmodule-object-outlook.md)

