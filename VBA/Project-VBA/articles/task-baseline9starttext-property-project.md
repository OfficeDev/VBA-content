---
title: Task.Baseline9StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline9StartText
ms.assetid: fc4280f5-69b1-627d-a894-c052de3be122
ms.date: 06/08/2017
---


# Task.Baseline9StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline9StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline9StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline9StartText** has any value, you should convert the value to a date for the **Baseline9Start** property.


