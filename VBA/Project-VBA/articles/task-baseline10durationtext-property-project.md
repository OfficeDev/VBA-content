---
title: Task.Baseline10DurationText Property (Project)
keywords: vbapj.chm131426
f1_keywords:
- vbapj.chm131426
ms.prod: project-server
api_name:
- Project.Task.Baseline10DurationText
ms.assetid: 4f7545f0-43e4-86ce-3665-8fca80ae9f4d
ms.date: 06/08/2017
---


# Task.Baseline10DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline10DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline10DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline10DurationText** has any value, you should convert the value to a date for the **TaskBaselineDuration** property.


