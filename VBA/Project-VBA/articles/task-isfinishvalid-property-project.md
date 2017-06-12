---
title: Task.IsFinishValid Property (Project)
ms.prod: project-server
ms.assetid: 13981c95-28fc-7b2f-d8b2-5b235bbe684e
ms.date: 06/08/2017
---


# Task.IsFinishValid Property (Project)

 **True** if the finish date of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.


## Syntax

 _expression_. **IsFinishValid**

 _expression_ An expression that returns a **Task** object.


## Remarks

The finish date of a manually scheduled task can be valid even though the start date and duration are invalid (empty).

To check the start date and duration, use the  **[IsStartValid](task-isstartvalid-property-project.md)** property and the **[IsDurationValid](task-isdurationvalid-property-project.md)** property.


## Property value

 **VARIANT**


