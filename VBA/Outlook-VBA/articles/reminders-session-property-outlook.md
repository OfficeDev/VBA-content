---
title: Reminders.Session Property (Outlook)
keywords: vbaol11.chm568
f1_keywords:
- vbaol11.chm568
ms.prod: outlook
api_name:
- Outlook.Reminders.Session
ms.assetid: 000e69b8-fd8c-1bd2-4cda-659faf210711
ms.date: 06/08/2017
---


# Reminders.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Reminders** object.


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


[Reminders Object](reminders-object-outlook.md)

