---
title: TaskItem.Read Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Read
ms.assetid: 88e5e300-e036-b511-905c-f0c238c97ade
ms.date: 06/08/2017
---


# TaskItem.Read Event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

 _expression_ . **Read**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

The  **Read** event differs from the **[Open](taskitem-open-event-outlook.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](inspector-object-outlook.md)** .


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

