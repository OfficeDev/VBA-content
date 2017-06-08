---
title: Task.OutlineLevel Property (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineLevel
ms.assetid: 7b852e27-bdbc-ee01-4146-c22b929adfa5
ms.date: 06/08/2017
---


# Task.OutlineLevel Property (Project)

Gets the level of the task in the outline hierarchy. Read/write  **Integer**.


## Syntax

 _expression_. **OutlineLevel**

 _expression_ A variable that represents a **Task** object.


## Remarks

A task with an outline level of 1 is at the highest level in the outline; there are no summary tasks above it. A task with an outline level of 3 has two summary tasks above it.


