---
title: Application.TaskRespectLinks Method (Project)
keywords: vbapj.chm2175
f1_keywords:
- vbapj.chm2175
ms.prod: project-server
api_name:
- Project.Application.TaskRespectLinks
ms.assetid: 1910b74a-7ea7-d0eb-97b9-aa79330952a0
ms.date: 06/08/2017
---


# Application.TaskRespectLinks Method (Project)

Moves one or more selected tasks so that their dates are determined by their task dependencies.


## Syntax

 _expression_. **TaskRespectLinks**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **TaskRespectLinks** method applies to both manually scheduled and automatically scheduled tasks. Changes in the start or finish dates of a task depend on the types of links for predecessor and successor tasks.

The  **TaskRespectLinks** method corresponds to the **Respect Links** command on the **Task** tab on the Ribbon.


## Example

Suppose a manually scheduled task has a predecessor task with a finish-to-start (FS) link, the finish date of the predecessor task is 7/15/2012, and the start date of the manually scheduled task is 7/20/2012 with no lag time. Running the  **TaskRespectLinks** method moves the start date back to 7/15/2012.


