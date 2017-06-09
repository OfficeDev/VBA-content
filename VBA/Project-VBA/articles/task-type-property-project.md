---
title: Task.Type Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Type
ms.assetid: 04a44733-c528-5887-113e-bdc70db8bb7a
ms.date: 06/08/2017
---


# Task.Type Property (Project)

Gets or sets the way the task is calculated; that is, which one of units, duration, or work are fixed. Read/write  **PjTaskFixedType**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **Type** property for a task can be one of the following **[PjTaskFixedType](pjtaskfixedtype-enumeration-project.md)** constants: **pjFixedUnits**, **pjFixedDuration**, or **pjFixedWork**. The default value is **pjFixedUnits** for both automatically scheduled and manually scheduled tasks. The default task type can be set with the **DefaultTaskType** property for the **Project** object, or on the **Schedule** tab in the **Project Options** dialog box.




 **Note**  Although the task type can be set for automatically scheduled tasks in the  **Task Information** dialog box, the **Task type** drop-down list is disabled for manually scheduled tasks. However, you can programmatically change the task type for manually scheduled tasks. The **Task.Type** property is read/write for both manually scheduled and automatically scheduled tasks.


