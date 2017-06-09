---
title: Task.RecalcFlags Property (Project)
ms.prod: project-server
api_name:
- Project.Task.RecalcFlags
ms.assetid: d5a5989e-b134-240b-fd37-11f4999e74bc
ms.date: 06/08/2017
---


# Task.RecalcFlags Property (Project)

Gets a bit mask, flagging one or more conditions that are driving the task. Read-only  **Long**.


## Syntax

 _expression_. **RecalcFlags**

 _expression_ A variable that represents a **Task** object.


## Remarks

Use the  **[PjRecalcDriverType](pjrecalcdrivertype-enumeration-project.md)** constants with the return value from the **RecalcFlags** property to determine which specific conditions are driving the task.


