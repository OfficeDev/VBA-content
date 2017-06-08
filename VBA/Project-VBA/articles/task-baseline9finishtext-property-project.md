---
title: Task.Baseline9FinishText Property (Project)
keywords: vbapj.chm131563
f1_keywords:
- vbapj.chm131563
ms.prod: project-server
api_name:
- Project.Task.Baseline9FinishText
ms.assetid: e12d7bdf-c7ff-092c-6907-3fe83d26daae
ms.date: 06/08/2017
---


# Task.Baseline9FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline9FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline9FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline9FinishText** has any value, you should convert the value to a date for the **Baseline9Finish** property.


