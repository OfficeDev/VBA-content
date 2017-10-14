---
title: Task.Baseline5DurationText Property (Project)
keywords: vbapj.chm131501
f1_keywords:
- vbapj.chm131501
ms.prod: project-server
api_name:
- Project.Task.Baseline5DurationText
ms.assetid: b6ac8444-0d82-2ff6-dad3-a982bc4413a2
ms.date: 06/08/2017
---


# Task.Baseline5DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline5DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline5DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline5DurationText** has any value, you should convert the value to a date for the **Baseline5Duration** property.


