---
title: Task.Status Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Status
ms.assetid: 4ea3a033-2306-8ae1-4e5e-c0420dcfa3dc
ms.date: 06/08/2017
---


# Task.Status Property (Project)

Gets the status of a specified task. Read-only  **PjStatusType**.


## Syntax

 _expression_. **Status**

 _expression_ A variable that represents a **Task** object.


## Remarks

The Status property can be one of the following  **[PjStatusType](pjstatustype-enumeration-project.md)** constants: **pjComplete**, **pjFutureTask**, **pjLate**, **pjNoData**, or **pjOnSchedule**.


