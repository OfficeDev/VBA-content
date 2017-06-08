---
title: Task.Baseline4FinishText Property (Project)
keywords: vbapj.chm131488
f1_keywords:
- vbapj.chm131488
ms.prod: project-server
api_name:
- Project.Task.Baseline4FinishText
ms.assetid: 9065f145-228b-5599-93fb-759da481a2a2
ms.date: 06/08/2017
---


# Task.Baseline4FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline4FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline4FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline4FinishText** has any value, you should convert the value to a date for the **Baseline4Finish** property.


