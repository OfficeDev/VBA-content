---
title: Task.Baseline8DurationText Property (Project)
keywords: vbapj.chm131546
f1_keywords:
- vbapj.chm131546
ms.prod: project-server
api_name:
- Project.Task.Baseline8DurationText
ms.assetid: a2410973-9a4a-d2b2-3a3b-610c23bb35b5
ms.date: 06/08/2017
---


# Task.Baseline8DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline8DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline8DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline8DurationText** has any value, you should convert the value to a date for the **Baseline8Duration** property.


