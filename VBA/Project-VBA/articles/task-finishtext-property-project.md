---
title: Task.FinishText Property (Project)
keywords: vbapj.chm132229
f1_keywords:
- vbapj.chm132229
ms.prod: project-server
api_name:
- Project.Task.FinishText
ms.assetid: 1dac5d15-30e3-060a-9c8a-98f7de556e3a
ms.date: 06/08/2017
---


# Task.FinishText Property (Project)

Gets or sets a string representation of the task finish date. Read/write  **String**.


## Syntax

 _expression_. **FinishText**

 _expression_ An expression that returns a **Task** object.


## Remarks

The  **FinishText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **FinishText** has any value, you should convert the value to a date for the **Finish** property.


