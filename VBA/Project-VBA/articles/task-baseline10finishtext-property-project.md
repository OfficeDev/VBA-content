---
title: Task.Baseline10FinishText Property (Project)
keywords: vbapj.chm131428
f1_keywords:
- vbapj.chm131428
ms.prod: project-server
api_name:
- Project.Task.Baseline10FinishText
ms.assetid: 1dde6265-6d9f-b7fd-8bc0-0f5315a6950e
ms.date: 06/08/2017
---


# Task.Baseline10FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline10FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline10FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline10FinishText** has any value, you should convert the value to a date for the **Baseline10Finish** property.


