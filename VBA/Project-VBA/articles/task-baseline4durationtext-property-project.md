---
title: Task.Baseline4DurationText Property (Project)
keywords: vbapj.chm131486
f1_keywords:
- vbapj.chm131486
ms.prod: project-server
api_name:
- Project.Task.Baseline4DurationText
ms.assetid: babe6ffe-b6e9-3bfd-a8d1-6384d8ab8a7e
ms.date: 06/08/2017
---


# Task.Baseline4DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline4DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline4DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline4DurationText** has any value, you should convert the value to a date for the **Baseline4Duration** property.


