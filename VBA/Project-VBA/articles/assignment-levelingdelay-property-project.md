---
title: Assignment.LevelingDelay Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.LevelingDelay
ms.assetid: b01087ec-9440-9288-3afe-6c0ed87e4a50
ms.date: 06/08/2017
---


# Assignment.LevelingDelay Property (Project)

Gets or sets the amount of time the assignment is delayed due to leveling. Read/write  **Variant**.


## Syntax

 _expression_. **LevelingDelay**

 _expression_ A variable that represents an **Assignment** object.


## Remarks

Project recalculates the leveling delay as resources are leveled across the project.

The  **LevelingDelay** property does not return any meaningful information for assignments of material resources. Setting a value returns a trappable error (error code 1101) when applied to assignments of material resources.


