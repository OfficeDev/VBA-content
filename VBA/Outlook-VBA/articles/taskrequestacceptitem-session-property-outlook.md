---
title: TaskRequestAcceptItem.Session Property (Outlook)
keywords: vbaol11.chm1772
f1_keywords:
- vbaol11.chm1772
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Session
ms.assetid: 5b50756f-1b1c-06d3-f3ee-24e71a66d01b
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


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


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

