---
title: CalendarModule.Session Property (Outlook)
keywords: vbaol11.chm2824
f1_keywords:
- vbaol11.chm2824
ms.prod: outlook
api_name:
- Outlook.CalendarModule.Session
ms.assetid: df23c975-9ac9-4ed9-0369-dce6b59e518a
ms.date: 06/08/2017
---


# CalendarModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **CalendarModule** object.


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


[CalendarModule Object](calendarmodule-object-outlook.md)

