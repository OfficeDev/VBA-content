---
title: TaskRequestItem.Session Property (Outlook)
keywords: vbaol11.chm1870
f1_keywords:
- vbaol11.chm1870
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Session
ms.assetid: a1206e37-cae8-3add-f679-70d5c7e7074c
ms.date: 06/08/2017
---


# TaskRequestItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **TaskRequestItem** object.


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


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

