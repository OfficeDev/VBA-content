---
title: Task.BaselineFinishText Property (Project)
ms.prod: project-server
api_name:
- Project.Task.BaselineFinishText
ms.assetid: 1cea31d3-ddc6-7fbc-ab40-8557c0790c40
ms.date: 06/08/2017
---


# Task.BaselineFinishText Property (Project)

Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.


## Syntax

 _expression_. **BaselineFinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **BaselineFinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **BaselineFinish** has any value, you should convert the value to a date for the **BaselineFinish** property.


