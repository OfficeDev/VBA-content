---
title: Task.Baseline5FinishText Property (Project)
keywords: vbapj.chm131503
f1_keywords:
- vbapj.chm131503
ms.prod: project-server
api_name:
- Project.Task.Baseline5FinishText
ms.assetid: 20ccdcac-c6b7-6728-3383-2c5bac33f60f
ms.date: 06/08/2017
---


# Task.Baseline5FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline5FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline5FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline5FinishText** has any value, you should convert the value to a date for the **Baseline5Finish** property.


