---
title: Task.DurationText Property (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Task.DurationText
ms.assetid: 4b0bbf0c-13fa-fcab-9940-b3471eb3509b
---


# Task.DurationText Property (Project)

Gets or sets a string representation of the task duration. Read/write  **String**.


## Syntax

 _expression_. **DurationText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **DurationText** has any value, you should convert the value to a date for the **Duration** property.


