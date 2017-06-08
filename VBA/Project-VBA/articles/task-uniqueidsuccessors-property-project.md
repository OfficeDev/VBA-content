---
title: Task.UniqueIDSuccessors Property (Project)
keywords: vbapj.chm132773
f1_keywords:
- vbapj.chm132773
ms.prod: project-server
api_name:
- Project.Task.UniqueIDSuccessors
ms.assetid: 2462e6da-8624-62f6-408e-0f50de82096d
ms.date: 06/08/2017
---


# Task.UniqueIDSuccessors Property (Project)

Gets or sets the unique identification ( **UniqueID** ) numbers of the successors of the task, separated by the list separator. Read/write **String**.


## Syntax

 _expression_. **UniqueIDSuccessors**

 _expression_ A variable that represents a **Task** object.


## Remarks

If a task has two successor tasks with the  **UniqueID** values of 10 and 12, for example, the **UniqueIDSuccessors** value is "10,12".


 **Note**   **UniqueID** values remain constant within a project and do not necessarily match the task **ID** values that can change with the position of the task in the outline or as tasks are deleted and added.


