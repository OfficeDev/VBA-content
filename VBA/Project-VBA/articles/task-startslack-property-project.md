---
title: Task.StartSlack Property (Project)
ms.prod: project-server
api_name:
- Project.Task.StartSlack
ms.assetid: 0a777363-9535-31b3-c24b-729a53b83190
ms.date: 06/08/2017
---


# Task.StartSlack Property (Project)

Gets the starting slack time of a task in minutes. Read-only  **Variant**.


## Syntax

 _expression_. **StartSlack**

 _expression_ A variable that represents a **Task** object.


## Remarks

Start slack is the difference between the early start and late start dates, where early start is the earliest date that a task can possibly start and late start is the latest date that a task can start without delaying the project finish date.


