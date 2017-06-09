---
title: Task.Baseline4StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline4StartText
ms.assetid: e4682921-053c-e93a-bcd6-ff77f4f3018a
ms.date: 06/08/2017
---


# Task.Baseline4StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline4StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline4StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline4StartText** has any value, you should convert the value to a date for the **Baseline4Start** property.


