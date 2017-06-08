---
title: CalendarView.Session Property (Outlook)
keywords: vbaol11.chm2609
f1_keywords:
- vbaol11.chm2609
ms.prod: outlook
api_name:
- Outlook.CalendarView.Session
ms.assetid: 550d9b8a-e980-9671-f45d-7ff54abdd591
ms.date: 06/08/2017
---


# CalendarView.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **CalendarView** object.


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


[CalendarView Object](calendarview-object-outlook.md)

