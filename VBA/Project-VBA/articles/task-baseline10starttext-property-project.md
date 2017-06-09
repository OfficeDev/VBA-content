---
title: Task.Baseline10StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline10StartText
ms.assetid: 1679422a-3bbe-ac70-48f6-bbcd3045e65c
ms.date: 06/08/2017
---


# Task.Baseline10StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline10StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline10StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline10StartText** has any value, you should convert the value to a date for the **Baseline10Start** property.


