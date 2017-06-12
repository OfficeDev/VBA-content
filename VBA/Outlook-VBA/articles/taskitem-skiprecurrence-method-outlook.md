---
title: TaskItem.SkipRecurrence Method (Outlook)
keywords: vbaol11.chm1756
f1_keywords:
- vbaol11.chm1756
ms.prod: outlook
api_name:
- Outlook.TaskItem.SkipRecurrence
ms.assetid: 19eb8a58-a13f-56ca-b742-a3780d8b0bf1
ms.date: 06/08/2017
---


# TaskItem.SkipRecurrence Method (Outlook)

Clears the current instance of a recurring task and sets the recurrence to the next instance of that task.


## Syntax

 _expression_ . **SkipRecurrence**

 _expression_ A variable that represents a **TaskItem** object.


### Return Value

 **False** indicates that the task was the last task in the recurrence, so there is no task to set the recurrence to. **True** indicates that the recurrence was successfully set to the next instance of that task.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

