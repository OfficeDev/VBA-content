---
title: Reminder.Session Property (Outlook)
keywords: vbaol11.chm556
f1_keywords:
- vbaol11.chm556
ms.prod: outlook
api_name:
- Outlook.Reminder.Session
ms.assetid: 30bd8c36-1afa-aae1-f050-47ad43af53f9
ms.date: 06/08/2017
---


# Reminder.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Reminder** object.


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


[Reminder Object](reminder-object-outlook.md)

