---
title: Task.Baseline6FinishText Property (Project)
keywords: vbapj.chm131518
f1_keywords:
- vbapj.chm131518
ms.prod: project-server
api_name:
- Project.Task.Baseline6FinishText
ms.assetid: 3c4c7ec1-6d73-e5c6-c097-9011eaebb371
ms.date: 06/08/2017
---


# Task.Baseline6FinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline6FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline6FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline6FinishText** has any value, you should convert the value to a date for the **Baseline6Finish** property.


