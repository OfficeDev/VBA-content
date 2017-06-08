---
title: Task.Baseline7FinishText Property (Project)
keywords: vbapj.chm131533
f1_keywords:
- vbapj.chm131533
ms.prod: project-server
api_name:
- Project.Task.Baseline7FinishText
ms.assetid: c6e180bc-12de-2fae-cb12-86c5ee25549d
ms.date: 06/08/2017
---


# Task.Baseline7FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline7FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline7FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline7FinishText** has any value, you should convert the value to a date for the **Baseline7Finish** property.


