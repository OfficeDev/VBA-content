---
title: PredecessorDrivers.TotalDetectedCount Property (Project)
ms.prod: project-server
api_name:
- Project.PredecessorDrivers.TotalDetectedCount
ms.assetid: 479cc962-5156-6f30-b304-5f4a6bc3abea
ms.date: 06/08/2017
---


# PredecessorDrivers.TotalDetectedCount Property (Project)

Gets the total number of predecessor tasks that affect the start date of a task. Read-only  **Long**.


## Syntax

 _expression_. **TotalDetectedCount**

 _expression_ A variable that represents a **PredecessorDrivers** object.


## Remarks

Predecessor tasks are tasks that are linked to the current task and occur before it. Predecessor tasks can have constraints and lag or lead time and can themselves have other predecessors that affect the total count of predecessor drivers.


## See also


#### Concepts


[PredecessorDrivers Collection Object](predecessordrivers-object-project.md)
