---
title: RecurrencePattern.Session Property (Outlook)
keywords: vbaol11.chm271
f1_keywords:
- vbaol11.chm271
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Session
ms.assetid: f30fce75-350c-6893-276a-47b19f211249
ms.date: 06/08/2017
---


# RecurrencePattern.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **RecurrencePattern** object.


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


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

