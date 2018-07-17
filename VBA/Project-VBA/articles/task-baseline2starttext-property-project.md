---
title: Task.Baseline2StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline2StartText
ms.assetid: b02c3892-73f2-59eb-25e9-7aa9bbe08a34
ms.date: 06/08/2017
---


# Task.Baseline2StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline2StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline2StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline2StartText** has any value, you should convert the value to a date for the **Baseline2Start** property.


