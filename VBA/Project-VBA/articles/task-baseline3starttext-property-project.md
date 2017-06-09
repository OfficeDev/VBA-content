---
title: Task.Baseline3StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline3StartText
ms.assetid: 1d9bfeb9-3272-aa45-4d9a-7c80cd842fee
ms.date: 06/08/2017
---


# Task.Baseline3StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline3StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline3StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline3StartText** has any value, you should convert the value to a date for the **Baseline3Start** property.


