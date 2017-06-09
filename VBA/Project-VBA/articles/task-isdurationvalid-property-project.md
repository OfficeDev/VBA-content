---
title: Task.IsDurationValid Property (Project)
ms.prod: project-server
ms.assetid: 303c5cab-b83a-37b6-c1da-207e91c45a86
ms.date: 06/08/2017
---


# Task.IsDurationValid Property (Project)

 **True** if the duration of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.


## Syntax

 _expression_. **IsDurationValid**

 _expression_ An expression that returns a **Task** object.


## Remarks

A manually scheduled task must have a valid start date and finish date for the duration to be valid.

To check the start date and finish date, use the  **[IsStartValid](task-isstartvalid-property-project.md)** property and the **[IsFinishValid](task-isfinishvalid-property-project.md)** property.


## Property value

 **VARIANT**


