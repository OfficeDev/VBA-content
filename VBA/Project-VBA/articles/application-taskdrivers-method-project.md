---
title: Application.TaskDrivers Method (Project)
keywords: vbapj.chm2279
f1_keywords:
- vbapj.chm2279
ms.prod: project-server
api_name:
- Project.Application.TaskDrivers
ms.assetid: 5c5e7563-e994-809b-7a9c-34f6ea338241
ms.date: 06/08/2017
---


# Application.TaskDrivers Method (Project)

Shows the  **Task Inspector** pane.


## Syntax

 _expression_. **TaskDrivers**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **TaskDrivers** method corresponds to the **Inspect Task** drop-down menu item on the **Task** tab of the Ribbon. The **TaskInspector** method has the same effect as the **[TaskInspector](application-taskinspector-method-project.md)** method.

The  **Task Inspector** pane includes factors that affect the task start and finish dates (task drivers such as calendars and predecessor tasks) and can also show warnings, suggestions, and ignored problems.


