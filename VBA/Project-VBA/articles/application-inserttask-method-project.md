---
title: Application.InsertTask Method (Project)
keywords: vbapj.chm2167
f1_keywords:
- vbapj.chm2167
ms.prod: project-server
api_name:
- Project.Application.InsertTask
ms.assetid: fe4676bf-8d9a-d6e9-2d5e-74fd047c3944
ms.date: 06/08/2017
---


# Application.InsertTask Method (Project)

Inserts a new task of the default mode above the selected task row or cell in a Gantt chart.


## Syntax

 _expression_. **InsertTask**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **InsertTask** method corresponds to the **Insert Task** command on the right-click option menu in a list of tasks. The new task is the default mode (manually or automatically scheduled), contains a task ID number, and the **Task Name** cell is selected with **<Type Task Name Here>**. Each task ID below the new row increases by one. 

By comparison, the  **[InsertBlankRow](application-insertblankrow-method-project.md)** method creates a blank row, where additional task information can be added programmatically. To create a manually scheduled task where the default mode is automatic, use the **[InsertManualTask](application-insertmanualtask-method-project.md)** method. To create an automatically scheduled task where the default mode is manual, use the **[InsertScheduledTask](application-insertscheduledtask-method-project.md)** method.


