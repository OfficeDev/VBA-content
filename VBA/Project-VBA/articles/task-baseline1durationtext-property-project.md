---
title: Task.Baseline1DurationText Property (Project)
keywords: vbapj.chm131441
f1_keywords:
- vbapj.chm131441
ms.prod: project-server
api_name:
- Project.Task.Baseline1DurationText
ms.assetid: 1fe64a4c-c4cd-8b18-6926-287789e3c30f
ms.date: 06/08/2017
---


# Task.Baseline1DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline1DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline1DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline1DurationText** has any value, you should convert the value to a date for the **Baseline1Duration** property.


