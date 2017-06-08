---
title: Project.ShowCriticalSlack Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ShowCriticalSlack
ms.assetid: fac1cf14-8f6f-34ca-7bab-71d444e78346
ms.date: 06/08/2017
---


# Project.ShowCriticalSlack Property (Project)

Gets or sets how much slack causes a task to be displayed as a critical task. Read/write  **Long**.


## Syntax

 _expression_. **ShowCriticalSlack**

 _expression_ A variable that represents a **Project** object.


## Remarks

If the slack time of a task does not exceed the number of days returned by the  **ShowCriticalSlack** property, Project displays the task as critical.


