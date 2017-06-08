---
title: Task.Baseline9DurationText Property (Project)
keywords: vbapj.chm131561
f1_keywords:
- vbapj.chm131561
ms.prod: project-server
api_name:
- Project.Task.Baseline9DurationText
ms.assetid: 8c9333e7-4b65-e317-4a9e-3d521de480ae
ms.date: 06/08/2017
---


# Task.Baseline9DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline9DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline9DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline9DurationText** has any value, you should convert the value to a date for the **Baseline9Duration** property.


