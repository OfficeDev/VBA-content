---
title: Application.TaskInspector Method (Project)
keywords: vbapj.chm1515
f1_keywords:
- vbapj.chm1515
ms.prod: project-server
api_name:
- Project.Application.TaskInspector
ms.assetid: cc2f34af-a4e0-8ad4-5dd1-9cf9663e342b
ms.date: 06/08/2017
---


# Application.TaskInspector Method (Project)

Displays the  **Task Inspector** pane.


## Syntax

 _expression_. **TaskInspector**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **TaskInspector** method corresponds to the **Inspect Task** drop-down menu item on the **TASK** ribbon. The **TaskInspector** method has the same effect as the **[TaskDrivers](application-taskdrivers-method-project.md)** method.

The  **Task Inspector** pane includes factors that affect the task start date and finish date (task drivers such as calendars and predecessor tasks) and can also show warnings, suggestions, and ignored problems.


