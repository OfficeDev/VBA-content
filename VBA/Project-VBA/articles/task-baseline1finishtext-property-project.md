---
title: Task.Baseline1FinishText Property (Project)
keywords: vbapj.chm131443
f1_keywords:
- vbapj.chm131443
ms.prod: project-server
api_name:
- Project.Task.Baseline1FinishText
ms.assetid: aa47b755-2670-a4e9-2c43-e6c90c625a06
ms.date: 06/08/2017
---


# Task.Baseline1FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline1FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline1FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline1FinishText** has any value, you should convert the value to a date for the **Baseline1Finish** property.


