---
title: Task.Baseline5StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline5StartText
ms.assetid: e2983aab-180e-0921-cadd-fdc9cd22908d
ms.date: 06/08/2017
---


# Task.Baseline5StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline5StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline5StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline5StartText** has any value, you should convert the value to a date for the **Baseline5Start** property.


