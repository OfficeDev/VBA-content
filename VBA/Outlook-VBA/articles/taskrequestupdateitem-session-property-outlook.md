---
title: TaskRequestUpdateItem.Session Property (Outlook)
keywords: vbaol11.chm1919
f1_keywords:
- vbaol11.chm1919
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Session
ms.assetid: 12e7fa2c-1067-4faa-c827-b1b1f8dc4238
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


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


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

