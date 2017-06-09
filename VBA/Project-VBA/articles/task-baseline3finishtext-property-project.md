---
title: Task.Baseline3FinishText Property (Project)
keywords: vbapj.chm131473
f1_keywords:
- vbapj.chm131473
ms.prod: project-server
api_name:
- Project.Task.Baseline3FinishText
ms.assetid: 126eecb3-bcfb-72c9-5da6-a54795b66f4d
ms.date: 06/08/2017
---


# Task.Baseline3FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline3FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline3FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline3FinishText** has any value, you should convert the value to a date for the **Baseline3Finish** property.


