---
title: Task.BaselineStartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.BaselineStartText
ms.assetid: cb50f6cd-eb28-24e2-862b-0963977bf815
ms.date: 06/08/2017
---


# Task.BaselineStartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **BaselineStartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **BaselineStartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **BaselineStartText** has any value, you should convert the value to a date for the **BaselineStart** property.


