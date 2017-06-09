---
title: Task.Baseline2DurationText Property (Project)
keywords: vbapj.chm131456
f1_keywords:
- vbapj.chm131456
ms.prod: project-server
api_name:
- Project.Task.Baseline2DurationText
ms.assetid: d0bacbcb-4976-451b-8b97-8bb70bb29c20
ms.date: 06/08/2017
---


# Task.Baseline2DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline2DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline2DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline2DurationText** has any value, you should convert the value to a date for the **Baseline2Duration** property.


