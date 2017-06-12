---
title: Task.Baseline7StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline7StartText
ms.assetid: 684af7b4-b7e5-bf33-1492-feb4004d6cad
ms.date: 06/08/2017
---


# Task.Baseline7StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline7StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline7StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline7StartText** has any value, you should convert the value to a date for the **Baseline7Start** property.


