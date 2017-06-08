---
title: Application.TaskDeliverableSync Method (Project)
keywords: vbapj.chm93
f1_keywords:
- vbapj.chm93
ms.prod: project-server
api_name:
- Project.Application.TaskDeliverableSync
ms.assetid: e5903c42-bade-959b-3c20-d02e3cf56b24
ms.date: 06/08/2017
---


# Application.TaskDeliverableSync Method (Project)

Synchronizes selected task deliverables in the active project with changes made in Project Web App. Available only in Project Professional.


## Syntax

 _expression_. **TaskDeliverableSync**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

If a deliverable is out of synchronization with Project Server, there is a red exclamation point to the left of the deliverable name in the  **Deliverables** pane.

The  **TaskDeliverableSync** method is equivalent to the **Sync Deliverables** command in the **Deliverable** drop-down menu on the **TASK** ribbon.


