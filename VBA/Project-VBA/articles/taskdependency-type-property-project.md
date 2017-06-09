---
title: TaskDependency.Type Property (Project)
ms.prod: project-server
api_name:
- Project.TaskDependency.Type
ms.assetid: fb8203b5-72ab-8b10-6698-461a75fce588
ms.date: 06/08/2017
---


# TaskDependency.Type Property (Project)

Gets or sets the link type of the task dependency. Read/write  **PjTaskLinkType**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **TaskDependency** object.


## Remarks

The task link types are sometimes abbreviated as FF (finish to finish), FS (finish to start), SF (start to finish), and SS (start to start).

The  **Type** property can be one of the following **[PjTaskLinkType](pjtasklinktype-enumeration-project.md)** constants: **pjFinishToFinish**, **pjFinishToStart**, **pjStartToFinish**, or **pjStartToStart**.


