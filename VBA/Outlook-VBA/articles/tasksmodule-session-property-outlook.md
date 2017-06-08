---
title: TasksModule.Session Property (Outlook)
keywords: vbaol11.chm2844
f1_keywords:
- vbaol11.chm2844
ms.prod: outlook
api_name:
- Outlook.TasksModule.Session
ms.assetid: 947b6795-21db-e2fb-b76b-43dc90520403
ms.date: 06/08/2017
---


# TasksModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **TasksModule** object.


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


[TasksModule Object](tasksmodule-object-outlook.md)

