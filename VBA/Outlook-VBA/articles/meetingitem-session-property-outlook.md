---
title: MeetingItem.Session Property (Outlook)
keywords: vbaol11.chm1400
f1_keywords:
- vbaol11.chm1400
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Session
ms.assetid: b18a448d-c3a6-e8cd-f251-30883e53e484
ms.date: 06/08/2017
---


# MeetingItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **MeetingItem** object.


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


[MeetingItem Object](meetingitem-object-outlook.md)

