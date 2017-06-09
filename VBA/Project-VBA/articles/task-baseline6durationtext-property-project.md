---
title: Task.Baseline6DurationText Property (Project)
keywords: vbapj.chm131516
f1_keywords:
- vbapj.chm131516
ms.prod: project-server
api_name:
- Project.Task.Baseline6DurationText
ms.assetid: b287077f-d296-eda0-45e1-8e5f25d096cd
ms.date: 06/08/2017
---


# Task.Baseline6DurationText Property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline6DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline6DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline6DurationText** has any value, you should convert the value to a date for the **Baseline6Duration** property.


