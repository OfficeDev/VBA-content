---
title: Task.Baseline6StartText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline6StartText
ms.assetid: fc304cc1-a90e-f9b8-d92f-81d8c9e27b66
ms.date: 06/08/2017
---


# Task.Baseline6StartText Property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

 _expression_. **Baseline6StartText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **Baseline6StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline6StartText** has any value, you should convert the value to a date for the **Baseline6Start** property.


