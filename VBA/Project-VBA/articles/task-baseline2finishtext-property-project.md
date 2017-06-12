---
title: Task.Baseline2FinishText Property (Project)
keywords: vbapj.chm131458
f1_keywords:
- vbapj.chm131458
ms.prod: project-server
api_name:
- Project.Task.Baseline2FinishText
ms.assetid: cfc6c6ba-9b23-13dd-1c25-74082fc69a0f
ms.date: 06/08/2017
---


# Task.Baseline2FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline2FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline2FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline2FinishText** has any value, you should convert the value to a date for the **Baseline2Finish** property.


