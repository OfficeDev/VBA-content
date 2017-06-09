---
title: Task.Baseline3DurationText Property (Project)
keywords: vbapj.chm131471
f1_keywords:
- vbapj.chm131471
ms.prod: project-server
api_name:
- Project.Task.Baseline3DurationText
ms.assetid: fa0ea4df-7658-5255-d91b-24a76005d7bc
ms.date: 06/08/2017
---


# Task.Baseline3DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline3DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline3DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline3DurationText** has any value, you should convert the value to a date for the **Baseline3Duration** property.


