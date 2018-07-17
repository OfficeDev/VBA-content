---
title: Task.ID Property (Project)
ms.prod: project-server
api_name:
- Project.Task.ID
ms.assetid: ce9b7773-77ae-c2ab-be11-08c20b57813e
ms.date: 06/08/2017
---


# Task.ID Property (Project)

Gets the identification number of a task. Read-only  **Long**.


## Syntax

 _expression_. **ID**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **ID** property changes when a task moves to a new location in a view such as the **Gantt Chart** or **Task Sheet**. Use the  **UniqueID** property if you want a constant reference to a task.


