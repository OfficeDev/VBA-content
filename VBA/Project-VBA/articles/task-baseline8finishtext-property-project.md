---
title: Task.Baseline8FinishText Property (Project)
keywords: vbapj.chm131548
f1_keywords:
- vbapj.chm131548
ms.prod: project-server
api_name:
- Project.Task.Baseline8FinishText
ms.assetid: 65704781-ed05-4127-ed76-8b3781c6bff3
ms.date: 06/08/2017
---


# Task.Baseline8FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline8FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline8FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline8FinishText** has any value, you should convert the value to a date for the **Baseline8Finish** property.


